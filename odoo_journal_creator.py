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


def safe_save_workbook(wb, file_path: Path):
    """Save workbook safely across platforms.

    Strategy: write to a sibling .tmp file first (never locked by Excel),
    then atomically replace the original.  This avoids writing to the
    locked file directly on Windows.

    - If os.replace() raises PermissionError (original locked by Excel):
        • Mac  — close just that workbook via AppleScript, then replace + reopen
        • Windows — try to close the workbook window, then replace + reopen.
          Only kills all Excel as a last resort if the replace still fails.
    - Mac first-save succeeds but Excel holds a stale in-memory copy:
        automatically closed & reopened so the user sees fresh data.
    """
    import subprocess
    import time
    import tempfile

    tmp_path = file_path.with_suffix(".tmp.xlsx")

    def _close_in_excel():
        """Ask Excel to close the specific workbook without saving."""
        fname = file_path.name
        if sys.platform == "darwin":
            subprocess.run(
                ["osascript", "-e",
                 f'tell application "Microsoft Excel"\n'
                 f'  set wb_list to every workbook whose name is "{fname}"\n'
                 f'  repeat with wb_item in wb_list\n'
                 f'    close wb_item saving no\n'
                 f'  end repeat\n'
                 f'end tell'],
                capture_output=True, timeout=5
            )
        elif os.name == "nt":
            # Try to close just the window with this filename in the title bar
            subprocess.run(
                ["taskkill", "/FI", f"WINDOWTITLE eq {fname}*", "/F"],
                capture_output=True
            )
            time.sleep(0.8)

    def _reopen():
        """Reopen the saved file so the user sees fresh data."""
        if sys.platform == "darwin":
            subprocess.run(["open", str(file_path)], capture_output=True)
        elif os.name == "nt":
            os.startfile(str(file_path))

    # ── Step 1: write to temp (always succeeds — Excel doesn't lock .tmp) ──
    try:
        wb.save(tmp_path)
    except Exception as e:
        print(f"❌ Could not write temp file: {e}")
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass
        return False

    # ── Step 2: atomic replace ──────────────────────────────────────────────
    try:
        os.replace(tmp_path, file_path)
        print(f"✅ Successfully saved '{file_path.name}'.")
        # Mac: Excel silently holds stale copy — close & reopen
        if sys.platform == "darwin":
            _close_in_excel()
            _reopen()
        return True

    except PermissionError:
        # File is locked — close it in Excel and retry replace
        print(f"⚠️ '{file_path.name}' is open in Excel. Closing to apply updates...")
        _close_in_excel()
        time.sleep(0.5)

        try:
            os.replace(tmp_path, file_path)
            print(f"✅ Successfully updated and saved '{file_path.name}'.")
            _reopen()
            return True
        except PermissionError:
            # Nuclear last resort on Windows only
            if os.name == "nt":
                print("⚠️ Still locked — force-closing all Excel instances...")
                subprocess.run(["taskkill", "/IM", "EXCEL.EXE", "/F"], capture_output=True)
                time.sleep(0.8)
                try:
                    os.replace(tmp_path, file_path)
                    print(f"✅ Saved '{file_path.name}' after closing Excel.")
                    _reopen()
                    return True
                except Exception as e2:
                    print(f"❌ Failed even after closing Excel: {e2}")
            else:
                print(f"❌ Failed to replace '{file_path.name}' — still locked.")
            try:
                tmp_path.unlink(missing_ok=True)
            except Exception:
                pass
            return False

    except Exception as e:
        print(f"❌ Replace error: {e}")
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass
        return False


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
            
            # Map column indices from row 3
            col_map = {str(ws.cell(row=3, column=c).value).strip().lower(): c for c in range(1, ws.max_column + 1) if ws.cell(row=3, column=c).value}
            c_status = col_map.get("status", 10)
            c_jstatus = col_map.get("journal status", 11)
            
            for item in selected_items:
                row_idx = item.get("row")
                if not row_idx:
                    continue
                
                created_edc = item.get("edc", False)
                created_ar = item.get("ar", False)
                
                new_status = ""
                if created_edc and created_ar:
                    new_status = "✅ Both"
                elif created_edc:
                    new_status = "✅ EDC"
                elif created_ar:
                    new_status = "✅ AR"
                else:
                    continue
                    
                current_val = str(ws.cell(row=row_idx, column=c_jstatus).value or "")
                
                # Merge status
                if "EDC" in current_val and created_ar:
                    new_status = "✅ Both"
                elif "AR" in current_val and created_edc:
                    new_status = "✅ Both"
                elif "Both" in current_val:
                    new_status = "✅ Both"
                
                ws.cell(row=row_idx, column=c_jstatus, value=new_status)
                
                # Keep original checkmark logic
                status9 = ws.cell(row=row_idx, column=c_status).value
                if status9 and "Match" in str(status9):
                    ws.cell(row=row_idx, column=c_status, value="✅ Journal Created")
                    
                updated += 1
                
            if updated > 0:
                safe_save_workbook(wb, recon_file)
                print(f"✅ Successfully updated {updated} rows in '{recon_file.name}' with Journal Status.")
    except Exception as e:
        print(f"⚠️ Failed to update status in reconciliation file: {e}")

def log_journal_creation(recon_file: Path, config_path: Path):
    """Append created journals to a master tracking log."""
    from openpyxl import load_workbook, Workbook
    import json
    from datetime import datetime

    log_file = OUTPUT_DIR / "journal_creation_log.xlsx"
    
    try:
        if not config_path or not config_path.exists():
            return

        with open(config_path, "r") as f:
            selected_items = json.load(f)
            
        wb_recon = load_workbook(recon_file, data_only=True)
        ws_recon = wb_recon["Daily Summary"]
        col_map = {str(ws_recon.cell(row=3, column=c).value).strip().lower(): c for c in range(1, ws_recon.max_column + 1) if ws_recon.cell(row=3, column=c).value}
        
        # Open or create log file
        if log_file.exists():
            wb_log = load_workbook(log_file)
            ws_log = wb_log.active
        else:
            wb_log = Workbook()
            ws_log = wb_log.active
            ws_log.title = "Journal Log"
            headers = ["Created At", "Original Recon File", "Bank", "Journal", "Transaction Date", "Merchant Amt", "Odoo Amt", "Created"]
            ws_log.append(headers)
            
        added = 0
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        for item in selected_items:
            row_idx = item.get("row")
            if not row_idx:
                continue
                
            created_edc = item.get("edc", False)
            created_ar = item.get("ar", False)
            
            if not created_edc and not created_ar:
                continue
                
            created_str = "Both" if (created_edc and created_ar) else ("EDC" if created_edc else "AR")
            
            bank = ws_recon.cell(row=row_idx, column=col_map.get("bank", 4)).value
            journal = ws_recon.cell(row=row_idx, column=col_map.get("journal", 5)).value
            date_val = ws_recon.cell(row=row_idx, column=col_map.get("date", 2)).value
            merch_amt = ws_recon.cell(row=row_idx, column=col_map.get("total bank", 6)).value
            odoo_amt = ws_recon.cell(row=row_idx, column=col_map.get("total odoo", 7)).value
            
            ws_log.append([
                now_str,
                recon_file.name,
                str(bank) if bank else "",
                str(journal) if journal else "",
                str(date_val) if date_val else "",
                merch_amt,
                odoo_amt,
                created_str
            ])
            added += 1
            
        if added > 0:
            # Protect the sheet from manual edits in Excel GUI
            ws_log.protection.sheet = True
            ws_log.protection.password = "ODOO_AUTO_SYSTEM_LOCK"
            safe_save_workbook(wb_log, log_file)
            print(f"📝 Appended {added} records to Journal Tracker Log ({log_file.name}).")
            
    except Exception as e:
        print(f"⚠️ Failed to write to journal tracker log: {e}")

def run_import_automation(import_file: Path, recon_file: Path, config_path: Path = None, email: str = "", password: str = "", headless: bool = False):
    """Launch Playwright, navigate to Odoo, and upload the import file."""
    if not ODOO_URL or not ODOO_DASHBOARD_URL or not ODOO_JOURNAL_IMPORT_URL:
        print("❌ ODOO_URL, ODOO_DASHBOARD_URL, or ODOO_JOURNAL_IMPORT_URL missing from .env")
        sys.exit(1)

    is_headless = headless or bool(email and password)

    import platform
    if platform.system() == "Windows":
        ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    else:
        ua = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

    with sync_playwright() as p:
        user_data_dir = os.path.abspath("playwright_profile")
        
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

        # Fix 404 Nginx Loop by upgrading HTTP to HTTPS automatically
        def upgrade_to_https(route, request):
            if request.url.startswith("http://"):
                secure_url = request.url.replace("http://", "https://", 1)
                route.fulfill(status=301, headers={"Location": secure_url})
            else:
                route.continue_()

        page.route("**/*", upgrade_to_https)

        print(f"🌐 Opening {ODOO_URL}")
        page.goto(ODOO_URL)

        # Auto-login if email/password provided and login form visible
        if email and password:
            try:
                page.wait_for_timeout(1000)
                if page.locator("input[name='login'], input#login").is_visible(timeout=3000):
                    print("🔑 Logging in to Odoo automatically...")
                    page.locator("input[name='login'], input#login").fill(email)
                    page.locator("input[name='password'], input#password").fill(password)
                    page.locator("button[type='submit']").click()
                    page.wait_for_timeout(2000)
            except Exception as e:
                print(f"ℹ️ Auto-login attempt note: {e}")

        print("\n⏳ Waiting for Odoo dashboard...")
        page.wait_for_function(
            "dashboardUrl => window.location.href.includes(dashboardUrl)",
            arg=ODOO_DASHBOARD_URL,
            timeout=0 if not is_headless else 30000
        )
        print("✅ Dashboard terdeteksi!")

        print(f"\n➡️ Mengarahkan ke form Import Jurnal: {ODOO_JOURNAL_IMPORT_URL}")
        page.goto(ODOO_JOURNAL_IMPORT_URL)
        
        page.wait_for_load_state("domcontentloaded")
        page.locator("xpath=/html/body/div[1]/div/div[1]/div/div[1]/div[1]/div[2]/span/span/button").wait_for(state="visible", timeout=30000)

        print(f"\n📤 Mengunggah file: {import_file.name} ...")
        xpath_upload_btn = "xpath=/html/body/div[1]/div/div[1]/div/div[1]/div[1]/div[2]/span/span/button"
        
        try:
            with page.expect_file_chooser(timeout=10000) as fc_info:
                page.locator(xpath_upload_btn).click()
            
            file_chooser = fc_info.value
            file_chooser.set_files(str(import_file.resolve()))
            print("✅ File uploaded successfully!")
        except Exception as e:
            print(f"❌ Failed to upload file. Ensure 'Upload File' button is visible.\nError: {e}")
            if not is_headless:
                input("\n[Press Enter to close browser]")
            context.close()
            return

        # Automate Test & Import
        print("🔍 Running Validation Test...")
        xpath_test_btn = "xpath=/html/body/div[1]/div/div[1]/div/div[1]/div[1]/div[2]/button[2]"
        xpath_valid_msg = "xpath=/html/body/div[1]/div/div[2]/div[2]/div/p"
        xpath_import_btn = "xpath=/html/body/div[1]/div/div[1]/div/div[1]/div[1]/div[2]/button[1]"
        
        try:
            page.locator(xpath_test_btn).wait_for(state="visible", timeout=30000)
            page.locator(xpath_test_btn).click()
            
            valid_locator = page.locator(xpath_valid_msg)
            valid_locator.wait_for(state="visible", timeout=15000)
            
            msg_text = valid_locator.inner_text()
            if "Everything seems valid" in msg_text or "valid" in msg_text.lower():
                print("✅ Validation successful: 'Everything seems valid.'")
                print("🚀 Executing Import...")
                page.locator(xpath_import_btn).click()
                
                try:
                    page.wait_for_url("**/web#action=*", timeout=15000)
                    print("✅ Successfully redirected to Journal Entries.")
                except Exception as e:
                    print(f"⚠️ Import submitted (timeout waiting for redirect: {e})")
                    
                # Always update status in Excel since import/creation is executed
                update_recon_file_status(recon_file, config_path)
                log_journal_creation(recon_file, config_path)
                    
                print("🎉 Import process finished!")
                page.wait_for_timeout(2000)
            else:
                print(f"⚠️ Odoo validation returned message: {msg_text}")
                # Still update recon file status as user submitted
                update_recon_file_status(recon_file, config_path)
                log_journal_creation(recon_file, config_path)
                if not is_headless:
                    input("\n[Press Enter in terminal to close browser]")
        except Exception as e:
            print(f"❌ Error during Test/Import: {e}")
            # Ensure recon file is updated
            update_recon_file_status(recon_file, config_path)
            log_journal_creation(recon_file, config_path)
            if not is_headless:
                input("\n[Press Enter in terminal to close browser]")
            
        context.close()


def main():
    import io as _io
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    elif hasattr(sys.stdout, "buffer"):
        sys.stdout = _io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="Automate Odoo Journal Creation via Import")
    parser.add_argument("--file", type=str, help="Path ke file Excel rekonsiliasi (opsional)")
    parser.add_argument("--config", type=str, help="Path ke file JSON konfigurasi jurnal (opsional)")
    parser.add_argument("--import-file", type=str, help="Path ke file import Excel yang sudah siap upload (opsional)")
    parser.add_argument("--email", type=str, default="", help="Odoo Email")
    parser.add_argument("--password", type=str, default="", help="Odoo Password")
    parser.add_argument("--headless", action="store_true", help="Run browser in headless mode")
    args = parser.parse_args()

    # 1. Get File
    if args.file:
        file_path = Path(args.file)
        if not file_path.exists():
            print(f"❌ File not found: {file_path}")
            sys.exit(1)
    else:
        file_path = get_latest_excel_file()
        if not file_path:
            sys.exit(1)
    
    print(f"📁 Processing data from: {file_path.name}")

    config_path = Path(args.config) if args.config else None

    # 2. Generate Import Excel or use existing
    if args.import_file:
        import_file = Path(args.import_file)
        if not import_file.exists():
            print(f"❌ Import file not found: {import_file}")
            sys.exit(1)
    else:
        import_file = generate_journal_import(file_path, config_path)
        if not import_file:
            sys.exit(1)

    # 3. Run Browser Automation
    run_import_automation(import_file, file_path, config_path, email=args.email, password=args.password, headless=args.headless)


if __name__ == "__main__":
    main()

