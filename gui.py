"""
Bank Reconciliation Studio — GUI Application
Modularized main application container.
"""
import sys
import os
import traceback
from pathlib import Path

# Suppress stdout/stderr if None (standard PyInstaller --noconsole behavior on Windows)
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")

if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

def _show_fatal_error(msg: str):
    print(f"\n[FATAL ERROR] {msg}", flush=True)
    try:
        log_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(".")
        (log_dir / "recon_crash.log").write_text(msg, encoding="utf-8")
        if sys.stderr:
            sys.stderr.write(msg + "\n")
            sys.stderr.flush()
    except Exception:
        pass
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Recon Studio - Fatal Error", f"Startup Error:\n\n{msg[:1500]}")
        root.destroy()
    except Exception:
        pass
    if sys.platform == "win32":
        try:
            input("\nPress Enter to exit...")
        except Exception:
            pass
    sys.exit(1)

try:
    import io
    import shutil
    import threading
    import subprocess
    from datetime import datetime
    import tkinter as tk
    from tkinter import filedialog, messagebox
    import customtkinter as ctk

    # Configure CustomTkinter
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    # Design System, Widgets, Views & Modals
    from ui.theme import (
        BASE_DIR, IS_WINDOWS, IS_MAC, FONT_FAMILY, BG, PANEL, SIDEBAR_BG, PREVIEW_BG,
        BORDER, BORDER_DARK, ACCENT, ACCENT_DARK, SUCCESS, ERROR, WARN, TEXT, MUTED, WHITE,
        init_fonts
    )
    from ui.widgets import _open_path, _maximize_window, _center_modal_on_parent
    from ui.views import SidebarView, LocalDashboardView, CloudDashboardView, LogConsoleView
    from ui.controllers import ReconController
    from ui.modals import (
        open_breakdown_modal, open_cleanup_modal, open_conflict_modal,
        open_sales_portal_modal, open_discrepancy_inspection_modal,
        open_manual_match_modal, open_journal_modal
    )
    from config import INPUT_DIR, MUTATION_DIR, OUTPUT_DIR, ODO_EXCEL_PATH, ODO_JOURNAL_EXCEL_PATH
    RECAP_DIR = BASE_DIR / "recap"

    init_fonts()
except Exception:
    _show_fatal_error(traceback.format_exc())




class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Bank Reconciliation Studio")
        self.configure(fg_color=BG)

        # Window sizing & centering
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        init_w = min(1320, max(1050, sw - 60))
        init_h = min(800, max(640, sh - 75))
        init_x = max(0, (sw - init_w) // 2)
        init_y = max(10, (sh - init_h) // 2 - 15)
        self.geometry(f"{init_w}x{init_h}+{init_x}+{init_y}")
        self.minsize(1020, 600)
        self.resizable(True, True)

        self._set_app_icon()
        self.controller = ReconController(self)
        self._last_output = None
        self._journal_window = None

        self._build_ui()
        self.update_idletasks()
        self.protocol("WM_DELETE_WINDOW", self._on_app_close)
        self._auto_scan_after_id = self.after(300, self._auto_scan_on_startup)

    def _set_app_icon(self):
        try:
            icon_ico = BASE_DIR / "assets" / "app_icon.ico"
            icon_png = BASE_DIR / "assets" / "app_icon.png"
            if IS_WINDOWS and icon_ico.exists():
                self.iconbitmap(str(icon_ico))
            elif icon_png.exists():
                from PIL import Image, ImageTk
                img = ImageTk.PhotoImage(Image.open(str(icon_png)))
                self.iconphoto(True, img)
                self._icon_ref = img
        except Exception:
            pass

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw - w)//2}+{max(10, (sh - h)//2 - 15)}")

    def _on_app_close(self):
        self.controller.stop_process()
        self.destroy()

    def _build_ui(self):
        # ── 1. Left Sidebar ──
        qa_handlers = {
            "open_output": self._open_output,
            "export_pdf": self._on_export_summary_pdf,
            "open_input": self._open_input,
            "open_mutation": self._open_mutation,
            "open_odoo_file": self._open_odoo_file,
            "open_journal_file": self._open_journal_file,
            "open_recap": self._open_recap,
        }

        self.sidebar = SidebarView(
            self,
            on_run=self._on_run,
            on_stop=self.controller.stop_process,
            on_manual_match=self._on_manual_match,
            on_sync_cloud=self._on_sync_cloud,
            on_journal=self._on_journal,
            on_company_change=self._on_company_changed,
            on_source_change=self._on_source_changed,
            quick_access_handlers=qa_handlers
        )
        self.sidebar.pack(side="left", fill="y")

        # ── 2. Main Content Area ──
        main_area = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        main_area.pack(side="left", fill="both", expand=True)

        # Top Header Bar
        topbar = ctk.CTkFrame(main_area, fg_color=SIDEBAR_BG, corner_radius=0, height=54)
        topbar.pack(fill="x")
        topbar.pack_propagate(False)

        tb_wrap = ctk.CTkFrame(topbar, fg_color="transparent")
        tb_wrap.pack(fill="both", expand=True, padx=20, pady=12)

        self.folder_label = ctk.CTkLabel(tb_wrap, text="", font=(FONT_FAMILY, 11, "bold"), text_color=SUCCESS, anchor="w")
        self.folder_label.pack(side="left", anchor="w")

        import config
        self.company_badge = ctk.CTkLabel(
            tb_wrap, text=f"🏢  {config.COMPANY_NAME}",
            font=(FONT_FAMILY, 10, "bold"), text_color=ACCENT,
            fg_color=PANEL, corner_radius=6, padx=10, height=28
        )
        self.company_badge.pack(side="right", anchor="e")

        ctk.CTkFrame(main_area, height=1, fg_color=BORDER).pack(fill="x")

        # Main Body Wrap
        card_wrap = ctk.CTkFrame(main_area, fg_color="transparent")
        card_wrap.pack(fill="both", expand=True, padx=16, pady=12)

        # Action Buttons Toolbar
        action_bar = ctk.CTkFrame(card_wrap, fg_color="transparent")
        action_bar.pack(fill="x", pady=(0, 8))

        def _sec_btn(text, cmd, color=TEXT, hover=PREVIEW_BG):
            return ctk.CTkButton(
                action_bar, text=text, height=34, fg_color=PANEL, hover_color=hover,
                border_color=BORDER_DARK, border_width=1, text_color=color,
                font=(FONT_FAMILY, 10, "bold"), corner_radius=6, command=cmd
            )

        self.upload_btn = _sec_btn("⬆ Upload", self._on_upload)
        self.upload_btn.pack(side="left", padx=(0, 6))

        self.sync_all_btn = _sec_btn("☁️ Sync to Cloud", self._on_sync_local_to_cloud, color=ACCENT)
        self.sync_all_btn.pack(side="left", padx=(0, 6))

        self.download_btn = _sec_btn("⬇ Download", self._on_download)
        self.download_btn.pack(side="left", padx=(0, 6))

        self.pdf_btn = _sec_btn("📑 Summary PDF", self._on_export_summary_pdf)
        self.pdf_btn.pack(side="left", padx=(0, 6))

        self.cleanse_btn = _sec_btn("🗑 Clean", self._on_cleanse, color=ERROR, hover="#FEE2E2")
        self.cleanse_btn.pack(side="left")

        # ── 3. Dual Dashboard Container ──
        self.dash_card = ctk.CTkFrame(card_wrap, fg_color=PANEL, corner_radius=10, border_color=BORDER, border_width=1)
        self.dash_card.pack(fill="x", pady=(0, 10))

        dash_hdr = ctk.CTkFrame(self.dash_card, fg_color="transparent")
        dash_hdr.pack(fill="x", padx=16, pady=(10, 6))

        # Segmented Tabs: Local vs Cloud
        self._active_tab = "local"
        tab_box = ctk.CTkFrame(dash_hdr, fg_color="transparent")
        tab_box.pack(side="left")

        self.tab_btn_local = ctk.CTkButton(
            tab_box, text="📁  Local Batch Session", height=30, width=145,
            fg_color=ACCENT, hover_color=ACCENT_DARK, text_color=WHITE,
            font=(FONT_FAMILY, 10, "bold"), corner_radius=6, command=lambda: self._switch_tab("local")
        )
        self.tab_btn_local.pack(side="left", padx=(0, 6))

        self.tab_btn_cloud = ctk.CTkButton(
            tab_box, text="☁️  Cloud Database", height=30, width=135,
            fg_color=WHITE, hover_color="#F1F5F9", border_color=BORDER_DARK, border_width=1,
            text_color=TEXT, font=(FONT_FAMILY, 10, "bold"), corner_radius=6, command=lambda: self._switch_tab("cloud")
        )
        self.tab_btn_cloud.pack(side="left")

        self.dash_refresh_btn = ctk.CTkButton(
            dash_hdr, text="↻ Refresh", height=26, width=70,
            fg_color="transparent", hover_color=PREVIEW_BG, border_color=BORDER_DARK, border_width=1,
            text_color=MUTED, font=(FONT_FAMILY, 10, "bold"), corner_radius=5, command=self.refresh_active_dashboard
        )
        self.dash_refresh_btn.pack(side="right", padx=(8, 0))

        self.dash_last_update = ctk.CTkLabel(dash_hdr, text="Updated: —", font=(FONT_FAMILY, 10, "bold"), text_color=MUTED)
        self.dash_last_update.pack(side="right")

        ctk.CTkFrame(self.dash_card, height=1, fg_color=BORDER).pack(fill="x", padx=16)

        # Tab Views
        clicks = {
            "open_input": self._open_input,
            "open_odoo_file": self._open_odoo_file,
            "open_mutation": self._open_mutation,
            "open_output": self._open_output,
            "on_journal": self._on_journal,
            "on_sync_cloud": self._on_sync_cloud,
            "show_breakdown": self._show_bank_breakdown_modal
        }
        self.local_dash = LocalDashboardView(self.dash_card, click_handlers=clicks)
        self.local_dash.pack(fill="x")

        self.cloud_dash = CloudDashboardView(self.dash_card, on_sync=self._on_sync_cloud)

        # ── 4. Log Console View ──
        self.console = LogConsoleView(card_wrap)
        self.console.pack(fill="both", expand=True)

        self.refresh_folder_status()

    # ── Logging & Status Delegation ──
    def log_write(self, text: str, tag: str = ""):
        self.console.write(text, tag)

    def set_status(self, text: str, color: str):
        self.console.set_status(text, color)

    def set_dates(self, min_d: str, max_d: str):
        try:
            d_from = datetime.strptime(min_d, "%Y-%m-%d").strftime("%d/%m/%Y")
            d_to = datetime.strptime(max_d, "%Y-%m-%d").strftime("%d/%m/%Y")
            self.sidebar.date_from_var.set(d_from)
            self.sidebar.date_to_var.set(d_to)
            self.log_write(f"\n📅 [Auto-Detect] Bank date range: {d_from} - {d_to}\n", "ok")
        except Exception as e:
            self.log_write(f"\n⚠️ Failed to log detected dates: {e}\n", "warn")

    def _switch_tab(self, tab: str):
        self._active_tab = tab
        if tab == "cloud":
            self.tab_btn_cloud.configure(fg_color=ACCENT, hover_color=ACCENT_DARK, text_color=WHITE, border_width=0)
            self.tab_btn_local.configure(fg_color=WHITE, hover_color="#F1F5F9", text_color=TEXT, border_color=BORDER_DARK, border_width=1)
            self.local_dash.pack_forget()
            self.console.pack_forget()
            self.dash_card.pack_configure(fill="both", expand=True)
            self.cloud_dash.pack(fill="both", expand=True)
            self.cloud_dash.update_summary()
        else:
            self.tab_btn_local.configure(fg_color=ACCENT, hover_color=ACCENT_DARK, text_color=WHITE, border_width=0)
            self.tab_btn_cloud.configure(fg_color=WHITE, hover_color="#F1F5F9", text_color=TEXT, border_color=BORDER_DARK, border_width=1)
            self.cloud_dash.pack_forget()
            self.dash_card.pack_configure(fill="x", expand=False)
            self.local_dash.pack(fill="x")
            self.console.pack(fill="both", expand=True)
            self.local_dash.update_summary(INPUT_DIR, MUTATION_DIR, OUTPUT_DIR, self.controller.is_running())


    def refresh_active_dashboard(self):
        if self._active_tab == "cloud":
            self.cloud_dash.update_summary()
        else:
            self.refresh_folder_status()
            self.local_dash.update_summary(INPUT_DIR, MUTATION_DIR, OUTPUT_DIR, self.controller.is_running())
        self.dash_last_update.configure(text=f"Updated: {datetime.now().strftime('%H:%M:%S')}")

    def refresh_folder_status(self):
        c_merch = sum(1 for f in INPUT_DIR.rglob("*") if f.is_file() and not f.name.startswith(".") and not f.name.startswith("~$")) if INPUT_DIR.exists() else 0
        c_mut = sum(1 for f in MUTATION_DIR.rglob("*.csv") if f.is_file()) if MUTATION_DIR.exists() else 0
        c_out = sum(1 for f in OUTPUT_DIR.glob("Reconciliation_*.xlsx") if f.is_file() and not f.name.startswith("~$")) if OUTPUT_DIR.exists() else 0
        has_p = 1 if ODO_EXCEL_PATH.exists() else 0
        has_j = 1 if ODO_JOURNAL_EXCEL_PATH.exists() else 0
        self.folder_label.configure(text=f"📂 Merchant: {c_merch}  |  📁 Mutation: {c_mut}  |  💳 Payment: {has_p}  |  📑 Journal: {has_j}  |  📊 Output: {c_out}")

    def _auto_scan_on_startup(self):
        self.refresh_active_dashboard()

    def on_process_started(self):
        self.sidebar.run_btn.configure(state="disabled", text="⏳ Running...")
        self.sidebar.stop_btn.pack(fill="x", pady=(6, 0))
        self.set_status("Running...", ACCENT)

    def on_process_ended(self):
        self.sidebar.run_btn.configure(state="normal", text="⚡  Run Reconciliation")
        self.sidebar.stop_btn.pack_forget()
        self.refresh_active_dashboard()

    # ── Action Handlers ──
    def _on_run(self):
        banks = self.sidebar.get_selected_banks()
        is_off = self.sidebar.offline_var.get()
        source = self.sidebar.recon_source_var.get()
        self.controller.run_reconciliation(banks, is_offline=is_off, source_mode=source)

    def _on_company_changed(self, ckey: str, cname: str):
        self.company_badge.configure(text=f"🏢  {cname}")
        self.log_write(f"\n🏢 Switched active company to: {cname} (.env.{ckey})\n", "ok")
        self.set_status(f"Company: {cname}", SUCCESS)
        self.refresh_active_dashboard()

    def _on_source_changed(self, mode: str):
        self.log_write(f"\n🔄 Data Source switched to: {'☁️ Cloud Database' if mode=='cloud' else '📁 Local Files'}\n", "dim")

    def _on_upload(self):
        from readers.file_detector import detect_file, copy_file
        files = filedialog.askopenfilenames(title="Select bank statement or mutation files", filetypes=[("All Supported", "*.xlsx;*.xls;*.csv;*.pdf;*.zip"), ("All Files", "*.*")])
        if not files: return
        try:
            self.log_write("\n── Uploading Files ──\n", "head")
            detections = [(Path(f), detect_file(Path(f))) for f in files]
            conflicts = []
            for path, res in detections:
                if res is not None:
                    dest_name = (path.name + ".zip") if res.wrap_as_zip else path.name
                    dest = res.target_dir / dest_name
                    if dest.exists():
                        conflicts.append((path, res, dest))

            action = open_conflict_modal(self, conflicts) if conflicts else "replace"
            copied = 0
            for path, res in detections:
                if res is None:
                    self.log_write(f"⚠️ Ignored: {path.name} (No pattern match)\n", "warn")
                    continue
                if any(p == path for p, _, _ in conflicts) and action == "skip":
                    continue
                copy_file(res, path)
                copied += 1
                self.log_write(f"✅ Uploaded: {path.name} → {res.target_dir.name}/\n", "ok")

            self.refresh_active_dashboard()
            self.set_status(f"Uploaded {copied} files", SUCCESS)

        except Exception as e:
            self.log_write(f"\n❌ Upload Error: {e}\n", "err")

    def _on_sync_local_to_cloud(self):
        from cloud_sync import is_cloud_configured, sync_local_to_cloud
        if not is_cloud_configured():
            self.log_write("\n❌ Cloud Sync Error: SUPABASE_URL not configured.\n", "err")
            self.set_status("Cloud sync not configured", ERROR)
            return

        def _bg():
            self.log_write("\n── ☁️ Syncing Local Files to Cloud Database ──\n", "head")
            self.set_status("Syncing to Cloud...", WARN)
            res = sync_local_to_cloud(user_profile=self.sidebar.email_var.get() or "Desktop-User")
            if res.get("success"):
                self.after(0, self.log_write, f"✅ Cloud Sync Complete: Synced {res.get('merchant_count',0)} merchant & {res.get('mutation_count',0)} mutation rows!\n", "ok")
                self.after(0, self.set_status, "Cloud Sync Successful", SUCCESS)
                self.after(0, self.refresh_active_dashboard)
            else:
                self.after(0, self.log_write, f"❌ Cloud Sync Failed: {res.get('error')}\n", "err")
                self.after(0, self.set_status, "Cloud Sync Failed", ERROR)

        threading.Thread(target=_bg, daemon=True).start()

    def _on_download(self):
        from odoo_downloader import download_via_xmlrpc
        d_from = self.sidebar.date_from_var.get().strip()
        d_to = self.sidebar.date_to_var.get().strip()
        banks = self.sidebar.get_selected_banks()

        if not d_from or not d_to:
            try:
                from main import scan_bank_dates_detailed
                scan_details = scan_bank_dates_detailed(banks)
                if scan_details.get("min_date") and scan_details.get("max_date"):
                    d_from = datetime.strptime(scan_details["min_date"], "%Y-%m-%d").strftime("%d/%m/%Y")
                    d_to = datetime.strptime(scan_details["max_date"], "%Y-%m-%d").strftime("%d/%m/%Y")
                    self.sidebar.date_from_var.set(d_from)
                    self.sidebar.date_to_var.set(d_to)
            except Exception:
                pass

        def _bg():
            self.log_write(f"\n── Downloading Odoo Payment & Journals ({d_from or 'Auto'} to {d_to or 'Auto'}) ──\n", "head")
            self.set_status("Downloading via XML-RPC...", WARN)
            ok = download_via_xmlrpc(date_from=d_from, date_to=d_to, banks=banks, mode="both")
            if ok:
                self.after(0, self.log_write, "✅ Odoo Download Finished!\n", "ok")
                self.after(0, self.set_status, "Download Complete", SUCCESS)
                self.after(0, self.refresh_active_dashboard)
            else:
                self.after(0, self.log_write, "❌ Download Failed!\n", "err")
                self.after(0, self.set_status, "Download Failed", ERROR)

        threading.Thread(target=_bg, daemon=True).start()


    def _on_cleanse(self):
        files_to_move = []
        for p in [INPUT_DIR, MUTATION_DIR, OUTPUT_DIR]:
            if p.exists():
                for f in p.rglob("*"):
                    if f.is_file() and not f.name.startswith(".") and not f.name.startswith("~$"):
                        files_to_move.append(f)
        if ODO_EXCEL_PATH.exists() and ODO_EXCEL_PATH not in files_to_move: files_to_move.append(ODO_EXCEL_PATH)
        if ODO_JOURNAL_EXCEL_PATH.exists() and ODO_JOURNAL_EXCEL_PATH not in files_to_move: files_to_move.append(ODO_JOURNAL_EXCEL_PATH)

        if not files_to_move:
            self.set_status("No files to clean", MUTED)
            return

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        target_recap = RECAP_DIR / ts

        def _do_clean():
            for f in files_to_move:
                try: rel = f.relative_to(BASE_DIR)
                except ValueError: rel = Path(f.name)
                dest = target_recap / rel
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(f), str(dest))
            self.log_write(f"✅ Archived {len(files_to_move)} files to: recap/{ts}/\n", "ok")
            self.set_status(f"Cleaned {len(files_to_move)} files", SUCCESS)
            self.refresh_active_dashboard()

        open_cleanup_modal(self, files_to_move, target_recap, _do_clean)

    def _on_export_summary_pdf(self):
        def _bg():
            self.set_status("Generating PDF Summary...", WARN)
            self.log_write("\n── Generating Executive PDF Summary ──\n", "head")
            try:
                from pdf_summary_generator import generate_executive_summary_pdf
                pdf_path = generate_executive_summary_pdf()
                self.after(0, self.log_write, f"✅ PDF Summary Generated: {pdf_path.name}\n", "ok")
                self.after(0, self.set_status, "PDF Ready", SUCCESS)
                _open_path(str(pdf_path))
            except Exception as e:
                self.after(0, self.log_write, f"❌ PDF Failed: {e}\n", "err")
                self.after(0, self.set_status, "PDF Error", ERROR)

        threading.Thread(target=_bg, daemon=True).start()

    # ── Modals Delegation ──
    def _show_bank_breakdown_modal(self):
        stats = getattr(self, "_latest_bank_stats", None) or getattr(getattr(self, "local_dash", None), "_latest_bank_stats", None)
        open_breakdown_modal(self, stats, on_open_output=self._open_output)


    def _on_sync_cloud(self):
        open_sales_portal_modal(self, output_dir=OUTPUT_DIR, email_var=self.sidebar.email_var, log_write_fn=self.log_write, set_status_fn=self.set_status)

    def _open_discrepancy_inspection_modal(self, item: dict, parent_win=None):
        open_discrepancy_inspection_modal(item, parent_win=parent_win or self)

    def _on_manual_match(self):
        open_manual_match_modal(self, output_dir=OUTPUT_DIR, log_write_fn=self.log_write, set_status_fn=self.set_status, on_open_journal=self._on_journal)

    def _on_journal(self):
        if getattr(self, "_journal_window", None) and self._journal_window.winfo_exists():
            try:
                self._journal_window.deiconify()
                self._journal_window.lift()
                self._journal_window.focus_force()
                return
            except Exception:
                pass
        if self.controller.is_running(): return
        self._journal_window = open_journal_modal(
            self, output_dir=OUTPUT_DIR, base_dir=BASE_DIR, venv_python=sys.executable,
            log_write_fn=self.log_write, set_status_fn=self.set_status,
            on_run_fn=self._on_run, on_done_fn=self.controller._on_done,
            on_manual_match_fn=self._on_manual_match, update_dashboard_fn=self.refresh_active_dashboard
        )

    # ── Folder & Quick Openers ──
    def _open_output(self):
        files = [f for f in OUTPUT_DIR.glob("Reconciliation_*.xlsx") if f.is_file() and not f.name.startswith("~$")] if OUTPUT_DIR.exists() else []
        if files: _open_path(str(max(files, key=lambda p: p.stat().st_mtime)))
        else: self.set_status("No reconciliation file found", WARN)

    def _open_odoo_file(self):
        if ODO_EXCEL_PATH.exists(): _open_path(str(ODO_EXCEL_PATH))
        elif ODO_JOURNAL_EXCEL_PATH.exists(): _open_path(str(ODO_JOURNAL_EXCEL_PATH))
        else: self._open_input()

    def _open_journal_file(self):
        if ODO_JOURNAL_EXCEL_PATH.exists(): _open_path(str(ODO_JOURNAL_EXCEL_PATH))
        else: self.set_status("Odoo journal file not found", WARN)

    def _open_input(self):
        INPUT_DIR.mkdir(exist_ok=True)
        _open_path(str(INPUT_DIR))

    def _open_mutation(self):
        MUTATION_DIR.mkdir(exist_ok=True)
        _open_path(str(MUTATION_DIR))

    def _open_recap(self):
        RECAP_DIR.mkdir(parents=True, exist_ok=True)
        _open_path(str(RECAP_DIR))


if __name__ == "__main__":
    try:
        if len(sys.argv) > 1:
            if sys.argv[1] == "--run-downloader":
                import odoo_downloader
                sys.argv = [sys.argv[0]] + sys.argv[2:]
                odoo_downloader.run_downloader()
                sys.exit(0)
            elif sys.argv[1] == "--run-main":
                import runpy
                runpy.run_module('main', run_name='__main__')
                sys.exit(0)

        app = App()
        app.mainloop()
    except Exception as e:
        import traceback
        _show_fatal_error(traceback.format_exc())

