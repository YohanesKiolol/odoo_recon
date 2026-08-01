"""
gui.py — Bank Reconciliation GUI. Cross-platform (Windows + Mac).

PyInstaller pattern:
  - GUI mode   : BankRekonsiliasi.exe            → opens this window
  - Worker mode: BankRekonsiliasi.exe --worker   → runs main logic, stdout only
"""
import sys
import os
import platform
from pathlib import Path

# ── Resolve base directory ────────────────────────────────────────────────────
BASE_DIR = (
    Path(sys.executable).parent   # frozen: next to the .exe
    if getattr(sys, "frozen", False)
    else Path(__file__).parent    # dev: next to gui.py
)
INPUT_DIR  = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
IS_WINDOWS = platform.system() == "Windows"

# ── Worker mode ───────────────────────────────────────────────────────────────
# The GUI re-launches the same binary with --worker for a clean stdout stream.
# Must be checked BEFORE any tkinter import so the headless worker doesn't
# require a display or GUI toolkit.
if "--worker" in sys.argv:
    sys.argv = [a for a in sys.argv if a != "--worker"]  # hide flag from main's argparse
    os.chdir(str(BASE_DIR))
    import runpy
    if getattr(sys, "frozen", False):
        main_path = str(Path(getattr(sys, "_MEIPASS", BASE_DIR)) / "main.py")
    else:
        main_path = str(BASE_DIR / "main.py")
    runpy.run_path(main_path, run_name="__main__")
    sys.exit(0)

# ── GUI-only imports (skipped in worker mode) ─────────────────────────────────
import threading
import subprocess
import shutil
import fnmatch
from datetime import datetime
import tkinter as tk
from tkinter import scrolledtext, filedialog
try:
    from tkcalendar import DateEntry
except ImportError:
    DateEntry = None

# Venv python path (dev mode only — frozen uses sys.executable)
if os.name == 'nt':
    _venv_python_path = BASE_DIR / ".venv" / "Scripts" / "python.exe"
else:
    _venv_python_path = BASE_DIR / ".venv" / "bin" / "python"

_venv_python = (
    sys.executable if getattr(sys, "frozen", False) 
    else str(_venv_python_path)
)

# ── Color palette ─────────────────────────────────────────────────────────────
BG          = "#1C2333"
PANEL       = "#252E42"
ACCENT      = "#3D8EF0"
ACCENT_DARK = "#2A6AC2"
SUCCESS     = "#27AE60"
ERROR       = "#E74C3C"
WARN        = "#F39C12"
TEXT        = "#E8EDF5"
MUTED       = "#8899AA"
WHITE       = "#FFFFFF"


def _open_path(path: str):
    """Open a file/folder in the OS default app."""
    if IS_WINDOWS:
        os.startfile(path) # type: ignore
    elif sys.platform == "darwin":
        subprocess.run(["open", path])
    else:
        subprocess.run(["xdg-open", path])


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bank Reconciliation Tool")
        self.configure(bg=BG)
        self.geometry("900x700")
        self.minsize(700, 500)
        self.resizable(True, True)
        self._running = False
        self._last_output: str | None = None
        self._build_ui()
        self._center()

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=PANEL, padx=30, pady=20)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🏦  Bank Reconciliation", bg=PANEL,
                 fg=WHITE, font=("Segoe UI", 20, "bold")).pack(anchor="w")
        tk.Label(hdr, text="Bandingkan transaksi Bank dengan Odoo secara otomatis",
                 bg=PANEL, fg=MUTED, font=("Segoe UI", 11)).pack(anchor="w", pady=(4, 0))

        # Input folder status
        folder_frame = tk.Frame(self, bg=BG, padx=30, pady=14)
        folder_frame.pack(fill="x")
        self._folder_label = tk.Label(folder_frame, bg=BG, fg=MUTED,
                                      font=("Segoe UI", 10), anchor="w")
        self._folder_label.pack(fill="x")
        self._refresh_folder_status()

        # Bank Selection
        bank_frame = tk.Frame(self, bg=BG, padx=30, pady=5)
        bank_frame.pack(fill="x")
        tk.Label(bank_frame, text="Pilih Bank:", bg=BG, fg=TEXT, font=("Segoe UI", 10)).pack(side="left", padx=(0, 10))

        self._bank_vars = {
            "BCA": tk.BooleanVar(value=True),
            "Mandiri": tk.BooleanVar(value=True),
            "BRI": tk.BooleanVar(value=True)
        }

        for b in ["BCA", "Mandiri", "BRI"]:
            cb = tk.Checkbutton(
                bank_frame, text=b, variable=self._bank_vars[b],
                bg=BG, fg=TEXT, selectcolor=PANEL, activebackground=BG, activeforeground=TEXT,
                font=("Segoe UI", 10), cursor="hand2"
            )
            cb.pack(side="left", padx=5)

        # ── Date Range Inputs ──
        date_frame = tk.Frame(self, bg=BG, padx=30)
        date_frame.pack(fill="x", pady=(10, 0))
        
        tk.Label(date_frame, text="Dari:", bg=BG, fg=TEXT, font=("Segoe UI", 9)).pack(side="left")
        if DateEntry:
            self._date_from_widget = DateEntry(date_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='mm/dd/yyyy')
            self._date_from_widget.pack(side="left", padx=(5, 15))
            self._date_from_widget.set_date(datetime.now().replace(day=1))
        else:
            self._date_from_var = tk.StringVar(value=datetime.now().strftime("%m/01/%Y"))
            tk.Entry(date_frame, textvariable=self._date_from_var, width=12).pack(side="left", padx=(5, 15))
        
        tk.Label(date_frame, text="Sampai:", bg=BG, fg=TEXT, font=("Segoe UI", 9)).pack(side="left")
        if DateEntry:
            self._date_to_widget = DateEntry(date_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='mm/dd/yyyy')
            self._date_to_widget.pack(side="left", padx=(5, 0))
            self._date_to_widget.set_date(datetime.now())
        else:
            self._date_to_var = tk.StringVar(value=datetime.now().strftime("%m/%d/%Y"))
            tk.Entry(date_frame, textvariable=self._date_to_var, width=12).pack(side="left", padx=(5, 0))

        # ── Action Toolbar (Single Row) ──
        action_frame = tk.Frame(self, bg=BG, padx=30)
        action_frame.pack(fill="x", pady=(5, 10))
        
        # Style sangat compact
        btn_style = {"font": ("Segoe UI", 9, "bold"), "relief": "flat", "cursor": "hand2", "padx": 12, "pady": 6}
        
        self._upload_btn = tk.Button(
            action_frame, text="📤 Upload",
            bg=PANEL, fg=TEXT, activebackground=ACCENT, activeforeground=WHITE,
            command=self._on_upload, **btn_style
        )
        self._upload_btn.pack(side="left", padx=(0, 6))

        self._download_btn = tk.Button(
            action_frame, text="📥 Download Odoo",
            bg=PANEL, fg=TEXT, activebackground=ACCENT, activeforeground=WHITE,
            command=self._on_download, **btn_style
        )
        self._download_btn.pack(side="left", padx=(0, 6))

        self._cleanse_btn = tk.Button(
            action_frame, text="🧹 Bersihkan",
            bg=WARN, fg=WHITE, activebackground="#E67E22", activeforeground=WHITE,
            command=self._on_cleanse, **btn_style
        )
        self._cleanse_btn.pack(side="left", padx=(0, 6))

        self._run_btn = tk.Button(
            action_frame, text="▶ Jalankan",
            bg=SUCCESS, fg=WHITE, activebackground="#219653", activeforeground=WHITE,
            command=self._on_run, **btn_style
        )
        self._run_btn.pack(side="left")

        # ── Utility Buttons ──
        util_frame = tk.Frame(self, bg=BG, padx=30)
        util_frame.pack(fill="x", pady=(0, 14))

        self._scan_btn = tk.Button(
            util_frame, text="🔍  Scan Data Saja",
            bg=PANEL, fg=TEXT, activebackground=ACCENT, activeforeground=WHITE,
            font=("Segoe UI", 9), relief="flat", cursor="hand2",
            padx=14, pady=4, command=self._on_scan,
        )
        self._scan_btn.pack(side="left", padx=(0, 10))

        self._open_input_btn = tk.Button(
            util_frame, text="📂  Buka File (Input)",
            bg=PANEL, fg=TEXT, activebackground=ACCENT, activeforeground=WHITE,
            font=("Segoe UI", 9), relief="flat", cursor="hand2",
            padx=14, pady=4, command=self._open_input,
        )
        self._open_input_btn.pack(side="left", padx=(0, 10))

        self._open_btn = tk.Button(
            util_frame, text="📁  Buka Hasil",
            bg=PANEL, fg=TEXT, activebackground=ACCENT, activeforeground=WHITE,
            font=("Segoe UI", 9), relief="flat", cursor="hand2",
            padx=14, pady=4, command=self._open_output, state="disabled",
        )
        self._open_btn.pack(side="left")

        # Status row
        self._status_var = tk.StringVar(value="Siap")
        status_row = tk.Frame(self, bg=BG, padx=30)
        status_row.pack(fill="x", pady=(0, 6))
        self._dot = tk.Label(status_row, text="●", bg=BG, fg=MUTED,
                             font=("Segoe UI", 13))
        self._dot.pack(side="left")
        tk.Label(status_row, textvariable=self._status_var, bg=BG, fg=MUTED,
                 font=("Segoe UI", 10)).pack(side="left", padx=(6, 0))

        # Log area
        log_frame = tk.Frame(self, bg=PANEL, padx=2, pady=2)
        log_frame.pack(fill="both", expand=True, padx=30, pady=(0, 20))
        self._log = scrolledtext.ScrolledText(
            log_frame, bg="#0F1623", fg=TEXT, insertbackground=TEXT,
            font=("Consolas", 11), relief="flat", state="disabled",
            width=90, height=28, wrap="word",
        )
        self._log.pack(fill="both", expand=True)
        self._log.tag_config("ok",   foreground=SUCCESS)
        self._log.tag_config("err",  foreground=ERROR)
        self._log.tag_config("warn", foreground=WARN)
        self._log.tag_config("head", foreground=ACCENT, font=("Consolas", 11, "bold"))
        self._log.tag_config("dim",  foreground=MUTED)

    def _center(self):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w,  h  = self.winfo_width(),       self.winfo_height()
        self.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")

    def _refresh_folder_status(self):
        files = list(INPUT_DIR.rglob("*")) if INPUT_DIR.exists() else []
        n = sum(1 for f in files if f.is_file())
        self._folder_label.config(
            text=f"📂  Input folder: {INPUT_DIR}   ({n} file ditemukan)",
            fg=SUCCESS if n > 0 else WARN,
        )

    def _log_write(self, text: str, tag: str = ""):
        self._log.config(state="normal")
        self._log.insert("end", text, tag)
        self._log.see("end")
        self._log.config(state="disabled")

    def _set_status(self, text: str, color: str):
        self._status_var.set(text)
        self._dot.config(fg=color)

    # ── New Feature Stubs ─────────────────────────────────────────────────────
    def _on_upload(self):
        try:
            from config import BCA_EXCEL_DIR, BCA_EXCEL_PATTERN, MANDIRI_ZIP_DIR, MANDIRI_ZIP_PATTERN, BRI_ZIP_DIR, BRI_ZIP_PATTERN, BRI_PDF_PATTERN, ODO_EXCEL_PATH
            
            files = filedialog.askopenfilenames(title="Pilih File Bank atau Odoo", filetypes=[("All Files", "*.*")])
            if not files:
                return
                
            self._log_write("\n── Mengupload File ──\n", "head")
            for f in files:
                path = Path(f)
                name = path.name.lower()
                
                target_dir = None
                if BCA_EXCEL_PATTERN.lower() in name:
                    target_dir = BCA_EXCEL_DIR
                elif fnmatch.fnmatch(name, MANDIRI_ZIP_PATTERN.lower()):
                    target_dir = MANDIRI_ZIP_DIR
                elif BRI_ZIP_PATTERN.lower() in name or BRI_PDF_PATTERN.lower() in name:
                    target_dir = BRI_ZIP_DIR
                elif "payments" in name and name.endswith(".xlsx"):
                    target_dir = ODO_EXCEL_PATH.parent
                
                if target_dir:
                    target_dir.mkdir(parents=True, exist_ok=True)
                    dest = target_dir / path.name
                    shutil.copy2(path, dest)
                    self._log_write(f"✅ Disalin: {path.name} -> {target_dir.name}/\n", "ok")
                else:
                    self._log_write(f"⚠️ Diabaikan: {path.name} (Tidak cocok pola bank)\n", "warn")
                    
            self._refresh_folder_status()
        except Exception as e:
            self._log_write(f"\n❌ Error saat Upload: {e}\nPastikan file .env Anda sudah terisi lengkap!\n", "error")

    def _on_download(self):
        if DateEntry:
            date_from = self._date_from_widget.get()
            date_to = self._date_to_widget.get()
        else:
            date_from = self._date_from_var.get().strip()
            date_to = self._date_to_var.get().strip()
        
        self._log_write(f"\n── Mendownload Odoo Payment (Dari {date_from} sampai {date_to}) ──\n", "head")
        self._set_status("Mendownload Odoo...", "orange")
        self._running = True
        
        def run():
            try:
                cmd = [
                    _venv_python, "odoo_downloader.py",
                    "--date-from", date_from,
                    "--date-to", date_to
                ]
                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                )
                
                for line in process.stdout:
                    self._log_write(line)
                    
                process.wait()
                if process.returncode == 0:
                    self._log_write("\n✅ Download Odoo Selesai!\n", "ok")
                else:
                    self._log_write(f"\n❌ Download Odoo gagal dengan kode {process.returncode}\n", "error")
                    
            except Exception as e:
                self._log_write(f"\n❌ Error: {e}\n", "error")
            finally:
                self._running = False
                self._refresh_folder_status()
                self._set_status("Siap", SUCCESS)
                
        threading.Thread(target=run, daemon=True).start()

    def _on_cleanse(self):
        try:
            from config import BCA_EXCEL_DIR, MANDIRI_ZIP_DIR, BRI_ZIP_DIR, ODO_EXCEL_PATH, OUTPUT_DIR
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            recap_dir = BASE_DIR / "recap" / timestamp
            
            self._log_write("\n── Membersihkan Data ──\n", "head")
            
            moved_count = 0
            dirs_to_clean = [BCA_EXCEL_DIR, MANDIRI_ZIP_DIR, BRI_ZIP_DIR, OUTPUT_DIR]
            
            for d in dirs_to_clean:
                if d.exists() and d.is_dir():
                    for f in d.glob("*"):
                        if f.is_file():
                            recap_dir.mkdir(parents=True, exist_ok=True)
                            shutil.move(str(f), str(recap_dir / f.name))
                            moved_count += 1
                            
            if ODO_EXCEL_PATH.exists():
                recap_dir.mkdir(parents=True, exist_ok=True)
                shutil.move(str(ODO_EXCEL_PATH), str(recap_dir / ODO_EXCEL_PATH.name))
                moved_count += 1
                
            if moved_count > 0:
                self._log_write(f"✅ {moved_count} file dipindahkan ke: recap/{timestamp}/\n", "ok")
            else:
                self._log_write("ℹ️ Tidak ada data yang perlu dibersihkan.\n", "dim")
                
            self._refresh_folder_status()
        except Exception as e:
            self._log_write(f"\n❌ Error saat Bersihkan Data: {e}\nPastikan file .env Anda sudah terisi lengkap!\n", "error")

    # ── Run / Scan ────────────────────────────────────────────────────────────
    def _on_scan(self):
        if self._running:
            return
        self._running = True
        selected_banks = [b.lower() for b, var in self._bank_vars.items() if var.get()]
        if not selected_banks:
            self._set_status("Pilih minimal 1 bank!", ERROR)
            self._running = False
            return

        self._scan_btn.config(state="disabled")
        self._run_btn.config(state="disabled")
        self._open_btn.config(state="disabled")
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")
        self._set_status("Scanning data...", WARN)
        self._refresh_folder_status()
        threading.Thread(target=self._run_script, args=(selected_banks, True), daemon=True).start()

    def _on_run(self):
        if self._running:
            return
        self._running = True
        selected_banks = [b.lower() for b, var in self._bank_vars.items() if var.get()]
        if not selected_banks:
            self._set_status("Pilih minimal 1 bank!", ERROR)
            self._running = False
            return

        self._scan_btn.config(state="disabled")
        self._run_btn.config(state="disabled", text="⏳  Berjalan...")
        self._open_btn.config(state="disabled")
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")
        self._set_status("Memproses...", WARN)
        self._refresh_folder_status()
        threading.Thread(target=self._run_script, args=(selected_banks, False), daemon=True).start()

    def _run_script(self, selected_banks, is_scan=False):
        # When frozen: re-launch the same .exe with --worker flag for clean stdout
        # When in dev:  launch main.py via the venv python
        if getattr(sys, "frozen", False):
            cmd = [sys.executable, "--worker"]
        elif _venv_python.exists():
            cmd = [str(_venv_python), str(BASE_DIR / "main.py")]
        else:
            cmd = [sys.executable, str(BASE_DIR / "main.py")]

        if selected_banks:
            cmd.extend(["--bank"] + selected_banks)
        
        if is_scan:
            cmd.append("--scan")

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
        )

        last_output_file = None
        if proc.stdout:
            for line in proc.stdout:
                ls = line.rstrip()
                if "reconciliation_" in ls and ".xlsx" in ls:
                    for part in ls.split():
                        if "reconciliation_" in part and ".xlsx" in part:
                            last_output_file = part.strip()
                self.after(0, self._log_write, ls + "\n", self._tag(ls))

        proc.wait()
        self.after(0, self._on_done, proc.returncode, last_output_file)

    @staticmethod
    def _tag(line: str) -> str:
        if "✅" in line:                          return "ok"
        if "❌" in line or "[ERROR]" in line:     return "err"
        if "⚠️" in line or "[WARN]" in line:      return "warn"
        if line.startswith("─") or line.startswith("="): return "head"
        if "[ODO]" in line or "skipped" in line.lower(): return "dim"
        return ""

    def _on_done(self, code: int, output_path: str | None):
        self._running = False
        self._scan_btn.config(state="normal")
        self._run_btn.config(state="normal", text="▶  Jalankan Rekonsiliasi")
        if code == 0:
            self._set_status("Selesai ✓", SUCCESS)
            self._open_btn.config(state="normal")
            self._last_output = output_path
            if output_path and Path(output_path).exists():
                _open_path(output_path)
        else:
            self._set_status("Gagal — lihat log di bawah", ERROR)

    def _open_output(self):
        path = self._last_output
        if path and Path(path).exists():
            _open_path(path)
        elif OUTPUT_DIR.exists():
            _open_path(str(OUTPUT_DIR))

    def _open_input(self):
        input_dir = BASE_DIR / "input"
        input_dir.mkdir(exist_ok=True)
        _open_path(str(input_dir))


if __name__ == "__main__":
    app = App()
    app.mainloop()
