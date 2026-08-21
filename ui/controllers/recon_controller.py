"""
Reconciliation & Background Process Controller.
Handles direct XML-RPC auto-recon threads, scanning, file uploads, cloud syncing, and terminal logging.
"""
import os
import sys
import subprocess
import threading
from pathlib import Path
from datetime import datetime

from ui.theme import BASE_DIR, IS_WINDOWS, SUCCESS, ERROR, WARN, MUTED
from ui.widgets import _open_path
from ui.modals import open_conflict_modal, open_cleanup_modal


class ReconController:
    def __init__(self, app):
        self.app = app
        self._running = False
        self._active_proc = None

    @staticmethod
    def tag_line(line: str) -> str:
        """Categorize log output line for syntax color-tagging."""
        if "✅" in line:                          return "ok"
        if "❌" in line or "[ERROR]" in line:     return "err"
        if "⚠️" in line or "[WARN]" in line:      return "warn"
        if line.startswith("─") or line.startswith("="): return "head"
        if "[ODO]" in line or "skipped" in line.lower(): return "dim"
        return ""

    def is_running(self) -> bool:
        return self._running

    def stop_process(self):
        """Terminate active background process safely."""
        proc = self._active_proc
        if proc:
            try:
                proc.terminate()
                self.app.after(500, lambda: proc.kill() if proc and proc.poll() is None else None)
            except Exception:
                pass
            self._active_proc = None
        self._running = False
        self.app.log_write("\n🛑 Process stopped by user.\n", "err")
        self.app.set_status("Process Stopped", ERROR)
        self.app.on_process_ended()

    def run_reconciliation(self, selected_banks: list[str], is_offline: bool = False, source_mode: str = "local"):
        """Run full reconciliation workflow with Direct XML-RPC or Offline engine."""
        if self._running:
            return

        if not selected_banks:
            self.app.set_status("Select at least 1 bank!", ERROR)
            return

        self._running = True
        self.app.on_process_started()

        def _bg():
            try:
                # ── Cloud Mode: Pre-download cloud data to local before running ──
                if source_mode == "cloud":
                    from cloud_sync import is_cloud_configured, sync_cloud_to_local
                    if not is_cloud_configured():
                        self.app.log_write("\n❌ Cloud reconciliation requires Supabase configuration in .env.\n", "err")
                        self.app.set_status("Cloud config missing", ERROR)
                        self.app.after(0, self._on_done, 1, None)
                        return

                    self.app.log_write("\n── ☁️ Fetching Transactions from Cloud Database ──\n", "head")
                    res = sync_cloud_to_local(banks=selected_banks)
                    if not res.get("success"):
                        self.app.log_write(f"\n❌ Failed to pull cloud data: {res.get('error')}\n", "err")
                        self.app.set_status("Cloud fetch failed", ERROR)
                        self.app.after(0, self._on_done, 1, None)
                        return
                    self.app.log_write(f"✅ Cloud data ready: {res.get('merchant_files', 0)} merchant file(s), {res.get('mutation_files', 0)} mutation file(s).\n", "ok")
                    self.app.refresh_folder_status()

                # Offline mode or standard runner
                if is_offline:
                    self.app.log_write("\n── Running Offline Reconciliation (Skipping Downloader) ──\n", "head")
                    self._run_main_engine(selected_banks)
                    return

                # Direct XML-RPC Auto-Recon
                self.app.log_write("\n── Starting Auto-Reconciliation (Direct XML-RPC API) ──\n", "head")
                from odoo_downloader import run_downloader, download_via_xmlrpc
                
                # Run worker command
                from config import BASE_DIR as CONFIG_BASE
                if getattr(sys, "frozen", False):
                    cmd = [sys.executable, "--run-downloader", "--mode", "auto_recon"]
                else:
                    from ui.theme import BASE_DIR
                    if os.name == 'nt':
                        vpy = BASE_DIR / ".venv" / "Scripts" / "python.exe"
                    else:
                        vpy = BASE_DIR / ".venv" / "bin" / "python"
                    cmd = [str(vpy), "odoo_downloader.py", "--mode", "auto_recon"]

                email = self.app.sidebar.email_var.get().strip()
                password = self.app.sidebar.password_var.get()
                if email and password:
                    cmd.extend(["--email", email, "--password", password])

                d_from = self.app.sidebar.date_from_var.get().strip()
                d_to = self.app.sidebar.date_to_var.get().strip()
                if d_from and d_to:
                    cmd.extend(["--date-from", d_from, "--date-to", d_to])

                if selected_banks:
                    cmd.extend(["--banks", ",".join(selected_banks)])

                env = os.environ.copy()
                env.pop("TCL_LIBRARY", None)
                env.pop("TK_LIBRARY", None)
                env["PYTHONIOENCODING"] = "utf-8"
                flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if IS_WINDOWS else 0
                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                    encoding="utf-8",
                    errors="replace",
                    creationflags=flags,
                    env=env
                )
                self._active_proc = process

                last_output_file = None
                if process.stdout:
                    for line in process.stdout:
                        ls = line.rstrip()
                        if ls.startswith("[DATE_RANGE]|"):
                            try:
                                _, min_d, max_d = ls.split("|")
                                self.app.set_dates(min_d, max_d)
                            except Exception:
                                pass
                            continue

                        ls_lower = ls.lower()
                        if "reconciliation_" in ls_lower and ".xlsx" in ls_lower:
                            for part in ls.split():
                                if "reconciliation_" in part.lower() and ".xlsx" in part.lower():
                                    last_output_file = part.strip()

                        self.app.after(0, self.app.log_write, ls + "\n", self.tag_line(ls))

                process.wait()
                if process.returncode != 0:
                    self.app.after(0, self.app.log_write, "\n❌ Auto-Recon Failed.\n", "err")
                    self.app.after(0, self._on_done, process.returncode, None)
                    return

                self.app.after(0, self.app.log_write, "\n✅ Auto-Recon Completed Successfully!\n", "ok")
                self.app.after(0, self._on_done, 0, last_output_file)

            except Exception as e:
                self.app.after(0, self.app.log_write, f"\n❌ Error during Auto-Recon: {e}\n", "err")
                self.app.after(0, self._on_done, 1, None)

        threading.Thread(target=_bg, daemon=True).start()

    def _run_main_engine(self, selected_banks: list[str]):
        """Run main.py directly."""
        try:
            from ui.theme import BASE_DIR
            if os.name == 'nt':
                vpy = BASE_DIR / ".venv" / "Scripts" / "python.exe"
            else:
                vpy = BASE_DIR / ".venv" / "bin" / "python"

            if getattr(sys, "frozen", False):
                cmd = [sys.executable, "--worker"]
            elif vpy.exists():
                cmd = [str(vpy), str(BASE_DIR / "main.py")]
            else:
                cmd = [sys.executable, str(BASE_DIR / "main.py")]

            if "all" in selected_banks:
                cmd.append("--all")
                selected_banks = [b for b in selected_banks if b != "all"]

            if selected_banks:
                cmd.extend(["--bank"] + selected_banks)

            cmd.append("--no-open")

            env = os.environ.copy()
            env.pop("TCL_LIBRARY", None)
            env.pop("TK_LIBRARY", None)
            env["PYTHONIOENCODING"] = "utf-8"
            flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if IS_WINDOWS else 0
            proc = subprocess.Popen(
                cmd,
                cwd=str(BASE_DIR),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=flags,
                env=env
            )
            self._active_proc = proc

            last_output_file = None
            if proc.stdout:
                for line in proc.stdout:
                    ls = line.rstrip()
                    if ls.startswith("[DATE_RANGE]|"):
                        try:
                            _, min_d, max_d = ls.split("|")
                            self.app.set_dates(min_d, max_d)
                        except Exception:
                            pass
                        continue

                    ls_lower = ls.lower()
                    if "reconciliation_" in ls_lower and ".xlsx" in ls_lower:
                        for part in ls.split():
                            if "reconciliation_" in part.lower() and ".xlsx" in part.lower():
                                last_output_file = part.strip()
                    self.app.after(0, self.app.log_write, ls + "\n", self.tag_line(ls))

            proc.wait()

            # Follow up with journal checker if report generated
            if proc.returncode == 0 and last_output_file:
                from config import ODO_JOURNAL_EXCEL_PATH
                if ODO_JOURNAL_EXCEL_PATH.exists():
                    self.app.after(0, self.app.log_write, "\n── Checking Journal Entries ──\n", "head")
                    j_cmd = [sys.executable, "journal_checker.py", last_output_file] if getattr(sys, "frozen", False) else [str(vpy), "journal_checker.py", last_output_file]
                    j_proc = subprocess.Popen(
                        j_cmd, cwd=str(BASE_DIR), stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                        text=True, encoding="utf-8", errors="replace", creationflags=flags, env=env
                    )
                    self._active_proc = j_proc
                    if j_proc.stdout:
                        for line in j_proc.stdout:
                            self.app.after(0, self.app.log_write, line)
                    j_proc.wait()

            self.app.after(0, self._on_done, proc.returncode, last_output_file)

        except Exception as e:
            self.app.after(0, self.app.log_write, f"\n❌ Engine Error: {e}\n", "err")
            self.app.after(0, self._on_done, 1, None)

    def _on_done(self, code: int, output_path: str | None):
        self._active_proc = None
        self._running = False
        self.app.on_process_ended()

        if code == 0:
            from config import OUTPUT_DIR
            target_file = None
            if output_path:
                p = Path(output_path)
                if p.exists():
                    target_file = p
                elif (OUTPUT_DIR / p.name).exists():
                    target_file = OUTPUT_DIR / p.name

            if target_file and target_file.exists():
                self.app.set_status("Finished ✓", SUCCESS)
                self.app._last_output = str(target_file)
                _open_path(str(target_file))
            else:
                self.app.set_status("Scan Complete ✓", SUCCESS)
        else:
            self.app.set_status("Failed — check logs below", ERROR)

        self.app.refresh_active_dashboard()
