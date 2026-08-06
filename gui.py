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
    # Force UTF-8 — Windows CP1252 can't encode ↔, ✅, ❌, ── etc.
    import io as _io
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    elif hasattr(sys.stdout, "buffer"):
        sys.stdout = _io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    import runpy
    if getattr(sys, "frozen", False):
        main_path = str(Path(getattr(sys, "_MEIPASS", BASE_DIR)) / "main.py")
    else:
        main_path = str(BASE_DIR / "main.py")
    runpy.run_path(main_path, run_name="__main__")
    sys.exit(0)

# ── GUI-only imports (skipped in worker mode) ─────────────────────────────────
if IS_WINDOWS:
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

import threading
import subprocess
import shutil
import fnmatch
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

try:
    from tkcalendar import DateEntry
except ImportError:
    DateEntry = None

# Configure CustomTkinter — light mode
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# Force PyInstaller to bundle dynamic imports
try:
    import config
    import odoo_downloader
except Exception:
    pass

# Venv python path (dev mode only — frozen uses sys.executable)
if os.name == 'nt':
    _venv_python_path = BASE_DIR / ".venv" / "Scripts" / "python.exe"
else:
    _venv_python_path = BASE_DIR / ".venv" / "bin" / "python"

_venv_python = (
    sys.executable if getattr(sys, "frozen", False)
    else str(_venv_python_path)
)

# ── Design System & Color Palette ─────────────────────────────────────────────
FONT_FAMILY = "Segoe UI" if IS_WINDOWS else "Helvetica Neue"
FONT_MONO   = "Consolas" if IS_WINDOWS else "Menlo"

# Theme Palette (Modern High-Contrast Clean Light)
BG          = "#F4F5F8"  # Cool neutral app background
PANEL       = "#FFFFFF"  # Pure white card background
SIDEBAR_BG  = "#FAFAFC"  # Soft off-white sidebar background
PREVIEW_BG  = "#F8FAFC"  # Table/Sub-card background
BORDER      = "#E2E8F0"  # Crisp subtle border color
BORDER_DARK = "#CBD5E1"  # Stronger border color for inputs

ACCENT      = "#6D28D9"  # Odoo Deep Violet
ACCENT_DARK = "#5B21B6"  # Hover Violet
SUCCESS     = "#059669"  # Vibrant Emerald
SUCCESS_DARK= "#047857"  # Hover Emerald
ERROR       = "#DC2626"  # Soft Crimson
ERROR_LIGHT = "#FEE2E2"
WARN        = "#D97706"  # Warm Amber
WARN_LIGHT  = "#FEF3C7"

TEXT        = "#0F172A"  # Deep slate text
MUTED       = "#64748B"  # Muted slate text
WHITE       = "#FFFFFF"

# CustomTkinter Color Mappings
CTK_FG      = ("#0F172A", "#0F172A")
CTK_ACCENT  = ACCENT
CTK_SUCCESS = SUCCESS
CTK_ERROR   = ERROR


def _open_path(path: str):
    """Open a file/folder in the OS default app."""
    if IS_WINDOWS:
        os.startfile(path) # type: ignore
    elif sys.platform == "darwin":
        subprocess.run(["open", path])
    else:
        subprocess.run(["xdg-open", path])


class CTkDateInput(ctk.CTkFrame):
    def __init__(self, master, variable=None, default_date=None, **kwargs):
        super().__init__(
            master, fg_color=WHITE, border_color=BORDER_DARK, border_width=1,
            corner_radius=6, height=34, **kwargs
        )
        self.pack_propagate(False)
        self._var = variable or tk.StringVar()
        
        self._cal_icon = tk.Label(
            self, text="📅", bg=WHITE, fg=MUTED,
            font=(FONT_FAMILY, 7), cursor="hand2"
        )
        self._cal_icon.pack(side="right", padx=(0, 4), pady=2)
        
        self._entry = ctk.CTkEntry(
            self, textvariable=self._var, placeholder_text="MM/DD/YYYY",
            height=28, border_width=0, fg_color="transparent", text_color=TEXT,
            font=(FONT_FAMILY, 10)
        )
        self._entry.pack(side="left", fill="x", expand=True, padx=(4, 0), pady=2)
        
        self._cal_icon.bind("<Button-1>", lambda e: self.open_calendar())
        self._entry.bind("<Button-1>", lambda e: self.open_calendar())
        self.bind("<Button-1>", lambda e: self.open_calendar())
        
        if default_date:
            self.set_date(default_date)
            
    def get(self):
        return self._var.get().strip()
        
    def set_date(self, d):
        s = d.strftime("%m/%d/%Y") if hasattr(d, "strftime") else str(d)
        self._var.set(s)
        self._entry.delete(0, "end")
        self._entry.insert(0, s)
            
    def open_calendar(self):
        try:
            import tkcalendar
            top = ctk.CTkToplevel(self)
            top.title("Select Date")
            top.geometry("260x220")
            top.configure(fg_color=PANEL)
            top.transient(self.winfo_toplevel())
            top.grab_set()
            
            x = self.winfo_rootx()
            y = self.winfo_rooty() + self.winfo_height() + 4
            top.geometry(f"+{x}+{y}")
            
            cal = tkcalendar.Calendar(
                top, selectmode="day", date_pattern="mm/dd/yyyy",
                background=ACCENT, foreground=WHITE, headersbackground=PREVIEW_BG,
                headersforeground=TEXT, selectbackground=ACCENT, selectforeground=WHITE,
                normalbackground=WHITE, normalforeground=TEXT
            )
            cal.pack(fill="both", expand=True, padx=8, pady=8)
            
            def _select():
                self.set_date(cal.get_date())
                top.destroy()
                
            cal.bind("<<CalendarSelected>>", lambda e: _select())
        except Exception:
            pass


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Bank Reconciliation Studio")
        self.configure(fg_color=BG)
        self.geometry("1100x740")
        self.minsize(900, 600)
        self.resizable(True, True)
        self._running = False
        self._last_output: str | None = None
        self._build_ui()
        self._center()

    # ── UI Construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        # DateEntry wrapper compatibility fallback
        class _DateVar:
            def __init__(self, var):       self._v = var
            def get(self):                  return self._v.get()
            def get_date(self):
                from datetime import datetime as _dt
                try:    return _dt.strptime(self._v.get(), "%m/%d/%Y")
                except: return _dt.now()
            def set_date(self, d):
                self._v.set(d.strftime("%m/%d/%Y") if hasattr(d, "strftime") else str(d))

        # Load bank logos from assets folder
        try:
            from PIL import Image as _PILImage
            assets_dir = (
                Path(getattr(sys, "_MEIPASS", BASE_DIR)) / "assets"
                if getattr(sys, "frozen", False)
                else BASE_DIR / "assets"
            )
            
            def _get_ctk_logo(filename, max_w=70, max_h=22):
                p = assets_dir / filename
                if not p.exists():
                    return None
                img = _PILImage.open(p).convert("RGBA")
                w_orig, h_orig = img.size
                scale = min(max_w / w_orig, max_h / h_orig)
                w = max(1, int(w_orig * scale))
                h = max(1, int(h_orig * scale))
                
                target_size = (w * 3, h * 3)
                img_scaled = img.resize(target_size, _PILImage.Resampling.LANCZOS)
                return ctk.CTkImage(light_image=img_scaled, dark_image=img_scaled, size=(w, h))

            self._logos = {
                "All":     _get_ctk_logo("all_logo.png"),
                "BCA":     _get_ctk_logo("bca_logo.png"),
                "Mandiri": _get_ctk_logo("mandiri_logo.png"),
                "BRI":     _get_ctk_logo("bri_logo.png"),
            }
            _has_logos = any(v is not None for v in self._logos.values())
        except Exception as e:
            print("LOGO LOAD ERROR:", e)
            self._logos = {}
            _has_logos = False

        # ── Sidebar Container ─────────────────────────────────────────────────
        sidebar_outer = ctk.CTkFrame(
            self, width=290, corner_radius=0,
            fg_color=SIDEBAR_BG, border_color=BORDER, border_width=1
        )
        sidebar_outer.pack(side="left", fill="y")
        sidebar_outer.pack_propagate(False)

        # Brand header — Pinned top
        brand = ctk.CTkFrame(sidebar_outer, fg_color="transparent", height=78)
        brand.pack(fill="x")
        brand.pack_propagate(False)

        brand_inner = ctk.CTkFrame(brand, fg_color="transparent")
        brand_inner.pack(fill="both", expand=True, padx=16, pady=14)

        ctk.CTkLabel(
            brand_inner, text="🏦 Bank Recon",
            font=(FONT_FAMILY, 18, "bold"), text_color=ACCENT,
            fg_color="transparent"
        ).pack(anchor="w")
        ctk.CTkLabel(
            brand_inner, text="Odoo Automation Tool",
            font=(FONT_FAMILY, 11), text_color=MUTED,
            fg_color="transparent"
        ).pack(anchor="w")
        
        ctk.CTkFrame(sidebar_outer, height=1, fg_color=BORDER).pack(fill="x")

        # Scrollable Sidebar Content
        scroll = ctk.CTkScrollableFrame(
            sidebar_outer, fg_color="transparent",
            scrollbar_button_color=BORDER,
            scrollbar_button_hover_color=MUTED
        )
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        # ── Credentials Section ───────────────────────────────────────────────
        sec_cred = ctk.CTkFrame(scroll, fg_color=PANEL, corner_radius=8, border_color=BORDER, border_width=1)
        sec_cred.pack(fill="x", padx=6, pady=(6, 8))
        
        ctk.CTkLabel(
            sec_cred, text="CREDENTIALS", font=(FONT_FAMILY, 9, "bold"),
            text_color=MUTED, fg_color="transparent"
        ).pack(anchor="w", padx=10, pady=(8, 4))

        ctk.CTkLabel(
            sec_cred, text="Email", font=(FONT_FAMILY, 10),
            text_color=TEXT, fg_color="transparent"
        ).pack(anchor="w", padx=10)
        
        self._email_var = tk.StringVar()
        ctk.CTkEntry(
            sec_cred, textvariable=self._email_var,
            placeholder_text="odoo@example.com",
            height=32, corner_radius=6, border_color=BORDER_DARK, fg_color=WHITE, text_color=TEXT,
            font=(FONT_FAMILY, 10)
        ).pack(fill="x", padx=10, pady=(2, 6))

        ctk.CTkLabel(
            sec_cred, text="Password", font=(FONT_FAMILY, 10),
            text_color=TEXT, fg_color="transparent"
        ).pack(anchor="w", padx=10)
        
        self._password_var = tk.StringVar()
        
        pass_frame = ctk.CTkFrame(sec_cred, fg_color=WHITE, border_color=BORDER_DARK, border_width=1, corner_radius=6, height=34)
        pass_frame.pack(fill="x", padx=10, pady=(2, 10))
        pass_frame.pack_propagate(False)

        self._show_pass = False
        def _toggle_pass():
            self._show_pass = not self._show_pass
            self._password_entry.configure(show="" if self._show_pass else "•")
            _btn_eye.configure(text="🙈" if self._show_pass else "👁")
            
        _btn_eye = tk.Label(
            pass_frame, text="👁", bg=WHITE, fg=MUTED,
            font=(FONT_FAMILY, 9), cursor="hand2"
        )
        _btn_eye.pack(side="right", padx=(4, 10), pady=4)
        _btn_eye.bind("<Button-1>", lambda e: _toggle_pass())
        
        self._password_entry = ctk.CTkEntry(
            pass_frame, textvariable=self._password_var,
            placeholder_text="Password", show="•",
            height=32, border_width=0, fg_color="transparent", text_color=TEXT,
            font=(FONT_FAMILY, 10)
        )
        self._password_entry.pack(side="left", fill="x", expand=True, padx=(8, 4))

        # ── Bank Target Section ───────────────────────────────────────────────
        sec_bank = ctk.CTkFrame(scroll, fg_color=PANEL, corner_radius=8, border_color=BORDER, border_width=1)
        sec_bank.pack(fill="x", padx=6, pady=4)

        ctk.CTkLabel(
            sec_bank, text="BANK TARGET", font=(FONT_FAMILY, 9, "bold"),
            text_color=MUTED, fg_color="transparent"
        ).pack(anchor="w", padx=10, pady=(8, 6))

        self._bank_vars = {
            "BCA":     tk.BooleanVar(value=False),
            "Mandiri": tk.BooleanVar(value=False),
            "BRI":     tk.BooleanVar(value=False),
            "All":     tk.BooleanVar(value=True),
        }
        _brand_bg = {"All": ACCENT, "BCA": "#0066AE", "Mandiri": "#F0A500", "BRI": "#004B87"}
        _brand_fg = {"All": WHITE, "BCA": WHITE, "Mandiri": "#1A1A2E", "BRI": WHITE}
        _bank_btns: dict = {}

        def _on_bank_toggle(name):
            if name == "All" and self._bank_vars["All"].get():
                for b in ["BCA", "Mandiri", "BRI"]:
                    self._bank_vars[b].set(False)
            elif name in ["BCA", "Mandiri", "BRI"] and self._bank_vars[name].get():
                self._bank_vars["All"].set(False)
            _refresh_bank_btns()

        def _refresh_bank_btns():
            for bname, btn in _bank_btns.items():
                sel = self._bank_vars[bname].get()
                btn.configure(
                    fg_color=WHITE if sel else PREVIEW_BG,
                    border_color=_brand_bg[bname] if sel else BORDER_DARK,
                    border_width=2 if sel else 1,
                )

        bank_grid = ctk.CTkFrame(sec_bank, fg_color="transparent")
        bank_grid.pack(fill="x", padx=8, pady=(0, 8))
        bank_grid.columnconfigure(0, weight=1)
        bank_grid.columnconfigure(1, weight=1)

        for bname, row, col in [("All", 0, 0), ("BCA", 0, 1),
                                  ("Mandiri", 1, 0), ("BRI", 1, 1)]:
            sel = self._bank_vars[bname].get()
            logo_img = self._logos.get(bname) if _has_logos else None

            def _make_cmd(n=bname):
                def _cmd():
                    self._bank_vars[n].set(not self._bank_vars[n].get())
                    _on_bank_toggle(n)
                return _cmd

            btn = ctk.CTkButton(
                bank_grid,
                text="" if logo_img else bname,
                image=logo_img,
                height=42, corner_radius=6,
                fg_color=WHITE if sel else PREVIEW_BG,
                hover_color=PREVIEW_BG,
                border_color=_brand_bg[bname] if sel else BORDER_DARK,
                border_width=2 if sel else 1,
                font=(FONT_FAMILY, 10, "bold"),
                command=_make_cmd(),
            )
            btn._logo_ref = logo_img
            btn.grid(row=row, column=col, padx=3, pady=3, sticky="ew")
            _bank_btns[bname] = btn
            
        _refresh_bank_btns()

        # ── Date Range Section ────────────────────────────────────────────────
        sec_date = ctk.CTkFrame(scroll, fg_color=PANEL, corner_radius=8, border_color=BORDER, border_width=1)
        sec_date.pack(fill="x", padx=6, pady=4)

        ctk.CTkLabel(
            sec_date, text="DATE RANGE", font=(FONT_FAMILY, 9, "bold"),
            text_color=MUTED, fg_color="transparent"
        ).pack(anchor="w", padx=10, pady=(8, 4))

        yesterday = datetime.now() - timedelta(days=1)
        import tkcalendar
        
        self._date_from_var = tk.StringVar()
        self._date_to_var = tk.StringVar()

        date_grid = ctk.CTkFrame(sec_date, fg_color="transparent")
        date_grid.pack(fill="x", padx=8, pady=(0, 8))
        date_grid.columnconfigure(0, weight=1)
        date_grid.columnconfigure(1, weight=1)
        
        ctk.CTkLabel(date_grid, text="From", font=(FONT_FAMILY, 9), text_color=MUTED,
                     fg_color="transparent").grid(row=0, column=0, sticky="w", padx=2)
        ctk.CTkLabel(date_grid, text="To", font=(FONT_FAMILY, 9), text_color=MUTED,
                     fg_color="transparent").grid(row=0, column=1, sticky="w", padx=2)
                     
        self._date_from_widget = CTkDateInput(
            date_grid, variable=self._date_from_var, default_date=yesterday
        )
        self._date_from_widget.grid(row=1, column=0, sticky="ew", padx=(0, 3))
        
        self._date_to_widget = CTkDateInput(
            date_grid, variable=self._date_to_var, default_date=yesterday
        )
        self._date_to_widget.grid(row=1, column=1, sticky="ew", padx=(3, 0))

        # ── Quick Folder Links Section ────────────────────────────────────────
        sec_links = ctk.CTkFrame(scroll, fg_color=PANEL, corner_radius=8, border_color=BORDER, border_width=1)
        sec_links.pack(fill="x", padx=6, pady=4)

        ctk.CTkLabel(
            sec_links, text="QUICK ACCESS", font=(FONT_FAMILY, 9, "bold"),
            text_color=MUTED, fg_color="transparent"
        ).pack(anchor="w", padx=10, pady=(6, 2))

        def _link_btn(parent, text, cmd):
            b = ctk.CTkButton(
                parent, text=text, height=28,
                fg_color="transparent", hover_color=PREVIEW_BG,
                text_color=ACCENT, font=(FONT_FAMILY, 10, "bold"),
                anchor="w", command=cmd
            )
            b.pack(fill="x", padx=6, pady=1)
            return b

        _link_btn(sec_links, "📂 Open Merchant", self._open_input)
        _link_btn(sec_links, "📁 Open Mutation", self._open_mutation)
        self._open_btn_sidebar = _link_btn(sec_links, "📊 Open Result", self._open_output)

        # ── Primary CTA Action Stack — Pinned Bottom ─────────────────────────
        ctk.CTkFrame(sidebar_outer, height=1, fg_color=BORDER).pack(fill="x")
        _cta = ctk.CTkFrame(sidebar_outer, fg_color="transparent")
        _cta.pack(fill="x", padx=12, pady=12)

        self._run_btn = ctk.CTkButton(
            _cta, text="▶  Reconciliation",
            height=40, fg_color=SUCCESS, hover_color=SUCCESS_DARK,
            text_color=WHITE, font=(FONT_FAMILY, 11, "bold"),
            corner_radius=8, command=self._on_run
        )
        self._run_btn.pack(fill="x", pady=(0, 6))

        self._journal_btn = ctk.CTkButton(
            _cta, text="📋  Generate Journal",
            height=40, fg_color=ACCENT, hover_color=ACCENT_DARK,
            text_color=WHITE, font=(FONT_FAMILY, 11, "bold"),
            corner_radius=8, command=self._on_journal
        )
        self._journal_btn.pack(fill="x")

        # ── Main Content Area ─────────────────────────────────────────────────
        main_area = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        main_area.pack(side="left", fill="both", expand=True)

        # ── Top Header Toolbar (Input Folder Status) ──────────────────────────
        topbar = ctk.CTkFrame(main_area, fg_color=SIDEBAR_BG, corner_radius=0, height=78)
        topbar.pack(fill="x")
        topbar.pack_propagate(False)
        
        tb_wrap = ctk.CTkFrame(topbar, fg_color="transparent")
        tb_wrap.pack(fill="both", expand=True, padx=20, pady=22)

        self._folder_label = ctk.CTkLabel(
            tb_wrap, text="",
            font=(FONT_FAMILY, 11, "bold"), text_color=SUCCESS,
            fg_color="transparent", anchor="w"
        )
        self._folder_label.pack(side="left", anchor="w")
        self._refresh_folder_status()

        ctk.CTkFrame(main_area, height=1, fg_color=BORDER).pack(fill="x")

        # ── Action Buttons & Console Log Area ─────────────────────────────────
        card_wrap = ctk.CTkFrame(main_area, fg_color="transparent")
        card_wrap.pack(fill="both", expand=True, padx=16, pady=14)

        # Action Buttons Toolbar (positioned directly on top of Console Log)
        action_bar = ctk.CTkFrame(card_wrap, fg_color="transparent")
        action_bar.pack(fill="x", pady=(0, 10))

        def _sec_btn(parent, text, cmd, color=TEXT, hover=PREVIEW_BG):
            return ctk.CTkButton(
                parent, text=text, height=34,
                fg_color=PANEL, hover_color=hover,
                border_color=BORDER_DARK, border_width=1,
                text_color=color, font=(FONT_FAMILY, 10, "bold"),
                corner_radius=6, command=cmd
            )

        self._upload_btn   = _sec_btn(action_bar, "⬆ Upload", self._on_upload)
        self._upload_btn.pack(side="left", padx=(0, 6))
        self._download_btn = _sec_btn(action_bar, "⬇ Download", self._on_download)
        self._download_btn.pack(side="left", padx=(0, 6))
        self._scan_btn     = _sec_btn(action_bar, "🔍 Scan", self._on_scan)
        self._scan_btn.pack(side="left", padx=(0, 6))
        self._cleanse_btn  = _sec_btn(action_bar, "🗑 Clean", self._on_cleanse, color=ERROR, hover=ERROR_LIGHT)
        self._cleanse_btn.pack(side="left")

        # Console Log Card
        log_card = ctk.CTkFrame(card_wrap, fg_color=PANEL, corner_radius=10, border_color=BORDER, border_width=1)
        log_card.pack(fill="both", expand=True)

        log_hdr = ctk.CTkFrame(log_card, fg_color="transparent")
        log_hdr.pack(fill="x", padx=16, pady=(12, 6))
        
        ctk.CTkLabel(
            log_hdr, text="Console Log",
            font=(FONT_FAMILY, 12, "bold"), text_color=TEXT,
            fg_color="transparent"
        ).pack(side="left")

        self._clear_log_btn = ctk.CTkButton(
            log_hdr, text="↻ Clear Log", height=26, width=70,
            fg_color="transparent", hover_color=PREVIEW_BG,
            border_color=BORDER_DARK, border_width=1,
            text_color=MUTED, font=(FONT_FAMILY, 10),
            corner_radius=5, command=self._clear_log
        )
        self._clear_log_btn.pack(side="right")
        
        ctk.CTkFrame(log_card, height=1, fg_color=BORDER).pack(fill="x", padx=16)

        log_frame = ctk.CTkFrame(log_card, fg_color="transparent")
        log_frame.pack(fill="both", expand=True, padx=10, pady=(6, 0))

        self._log = tk.Text(
            log_frame, bg=PANEL, fg=TEXT,
            font=(FONT_MONO, 9), relief="flat", state="disabled",
            wrap="word", borderwidth=0, highlightthickness=0,
            padx=10, pady=10, insertbackground=TEXT
        )
        _log_scroll = tk.Scrollbar(log_frame, command=self._log.yview)
        self._log.configure(yscrollcommand=_log_scroll.set)
        _log_scroll.pack(side="right", fill="y")
        self._log.pack(fill="both", expand=True)

        self._log.tag_config("ok",   foreground=SUCCESS)
        self._log.tag_config("err",  foreground=ERROR)
        self._log.tag_config("warn", foreground=WARN)
        self._log.tag_config("head", foreground=ACCENT, font=(FONT_MONO, 9, "bold"))
        self._log.tag_config("dim",  foreground=MUTED)

        # ── Integrated Status Bar ─────────────────────────────────────────────
        statusbar = ctk.CTkFrame(log_card, fg_color=PREVIEW_BG, corner_radius=0, border_color=BORDER, border_width=1)
        statusbar.pack(fill="x", side="bottom")
        
        st = ctk.CTkFrame(statusbar, fg_color="transparent")
        st.pack(fill="x", padx=16, pady=6)
        
        self._dot = ctk.CTkLabel(
            st, text="●", width=14,
            font=(FONT_FAMILY, 12), text_color=SUCCESS,
            fg_color="transparent"
        )
        self._dot.pack(side="left")
        
        self._status_var = tk.StringVar(value="Ready")
        ctk.CTkLabel(
            st, textvariable=self._status_var,
            font=(FONT_FAMILY, 10, "bold"), text_color=TEXT,
            fg_color="transparent"
        ).pack(side="left", padx=(4, 0))

        # Alias for completion callback compatibility
        self._open_btn = self._open_btn_sidebar

    def _center(self):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w,  h  = self.winfo_width(),       self.winfo_height()
        self.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")

    def _refresh_folder_status(self):
        files = list(INPUT_DIR.rglob("*")) if INPUT_DIR.exists() else []
        n = sum(1 for f in files if f.is_file())
        color = SUCCESS if n > 0 else WARN
        self._folder_label.configure(
            text=f"📂 Input folder: {INPUT_DIR}   ({n} file{'s' if n != 1 else ''} found)",
            text_color=color,
        )

    def _log_write(self, text: str, tag: str = ""):
        self._log.config(state="normal")
        self._log.insert("end", text, tag)
        self._log.see("end")
        self._log.config(state="disabled")

    def _set_status(self, text: str, color: str):
        self._status_var.set(text)
        self._dot.configure(text_color=color)

    def _clear_log(self):
        self._log.config(state="normal")
        self._log.delete("1.0", tk.END)
        self._log.config(state="disabled")

    # ── Feature Operations ────────────────────────────────────────────────────
    def _on_upload(self):
        try:
            from config import (
                BCA_EXCEL_DIR, BCA_EXCEL_PATTERN, 
                MANDIRI_ZIP_DIR, MANDIRI_ZIP_PATTERN, 
                BRI_ZIP_DIR, BRI_PDF_PATTERN, ODO_EXCEL_PATH,
                MUTATION_DIR, BCA_MUTATION_PATTERN,
                MANDIRI_MUTATION_PATTERN, BRI_MUTATION_PATTERN
            )
            
            files = filedialog.askopenfilenames(title="Select Bank or Odoo File", filetypes=[("All Files", "*.*")])
            if not files:
                return
                
            self._log_write("\n── Uploading Files ──\n", "head")
            for f in files:
                path = Path(f)
                name = path.name.lower()
                
                target_dir = None
                matched_bank = None
                is_mutation = False
                
                # 1. Check Mutations First
                if BCA_MUTATION_PATTERN and BCA_MUTATION_PATTERN.lower() in name:
                    target_dir = MUTATION_DIR / "bca"
                    matched_bank = "bca"
                    is_mutation = True
                elif MANDIRI_MUTATION_PATTERN and MANDIRI_MUTATION_PATTERN.lower() in name:
                    target_dir = MUTATION_DIR / "mandiri"
                    matched_bank = "mandiri"
                    is_mutation = True
                elif BRI_MUTATION_PATTERN and BRI_MUTATION_PATTERN.lower() in name:
                    target_dir = MUTATION_DIR / "bri"
                    matched_bank = "bri"
                    is_mutation = True
                
                # 2. If not mutation, check Input (EDC)
                elif BCA_EXCEL_PATTERN and BCA_EXCEL_PATTERN.lower() in name:
                    target_dir = BCA_EXCEL_DIR
                    matched_bank = "bca"
                elif fnmatch.fnmatch(name, MANDIRI_ZIP_PATTERN.lower()):
                    target_dir = MANDIRI_ZIP_DIR
                    matched_bank = "mandiri"
                elif "payments" in name and name.endswith(".xlsx"):
                    target_dir = ODO_EXCEL_PATH.parent
                else:
                    from config import BANK_ACCOUNTS, BRI_ZIP_PATTERN
                    if name.endswith(".zip") or name.endswith(".pdf") or name.endswith(".csv"):
                        for alias, acc_info in BANK_ACCOUNTS.get("bri", {}).items():
                            mid = acc_info.get("mid", "")
                            if mid and mid.lower() in name:
                                target_dir = BRI_ZIP_DIR
                                matched_bank = "bri"
                                break
                        
                        if not matched_bank and BRI_ZIP_PATTERN and BRI_ZIP_PATTERN.lower() in name:
                            target_dir = BRI_ZIP_DIR
                            matched_bank = "bri"
                
                if matched_bank:
                    from config import BANK_ACCOUNTS
                    matched_alias = None
                    accounts = BANK_ACCOUNTS.get(matched_bank, {})
                    for alias, acc_info in accounts.items():
                        identifier = acc_info.get("acc", "") if is_mutation else acc_info.get("mid", "")
                        if identifier and identifier.lower() in name:
                            matched_alias = alias
                            break
                    
                    if not matched_alias:
                        if "main" in accounts:
                            matched_alias = "main"
                        elif accounts:
                            matched_alias = list(accounts.keys())[0]
                        else:
                            matched_alias = "main"
                    target_dir = target_dir / matched_alias

                if target_dir:
                    target_dir.mkdir(parents=True, exist_ok=True)
                    dest = target_dir / path.name
                    shutil.copy2(path, dest)
                    self._log_write(f"✅ Copied: {path.name} -> {target_dir.parent.name}/{target_dir.name}/\n", "ok")
                else:
                    self._log_write(f"⚠️ Ignored: {path.name} (No bank pattern matched)\n", "warn")
                    
            self._refresh_folder_status()
        except Exception as e:
            self._log_write(f"\n❌ Error during Upload: {e}\nEnsure your .env file is configured correctly!\n", "err")

    def _on_download(self):
        if DateEntry:
            date_from = self._date_from_widget.get()
            date_to = self._date_to_widget.get()
        else:
            date_from = self._date_from_var.get().strip()
            date_to = self._date_to_var.get().strip()
        
        self._log_write(f"\n── Downloading Odoo Payment (From {date_from} to {date_to}) ──\n", "head")
        self._set_status("Downloading Odoo...", WARN)
        self._running = True
        
        def run():
            try:
                if getattr(sys, "frozen", False):
                    cmd = [sys.executable, "--run-downloader", "--date-from", date_from, "--date-to", date_to]
                else:
                    cmd = [_venv_python, "odoo_downloader.py", "--date-from", date_from, "--date-to", date_to]
                    
                email = self._email_var.get().strip()
                password = self._password_var.get()
                if email and password:
                    cmd.extend(["--email", email, "--password", password])
                    
                selected_banks = [b for b, var in self._bank_vars.items() if var.get()]
                if selected_banks:
                    cmd.extend(["--banks", ",".join(selected_banks)])

                env = os.environ.copy()
                env.pop("TCL_LIBRARY", None)
                env.pop("TK_LIBRARY", None)
                env["PYTHONIOENCODING"] = "utf-8"
                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                    encoding="utf-8",
                    errors="replace",
                    env=env
                )
                
                for line in process.stdout:
                    self._log_write(line)
                    
                process.wait()
                if process.returncode == 0:
                    self._log_write("\n✅ Odoo Download Finished!\n", "ok")
                else:
                    self._log_write(f"\n❌ Odoo Download failed with code {process.returncode}\n", "err")
                    
            except Exception as e:
                self._log_write(f"\n❌ Error: {e}\n", "err")
            finally:
                self._running = False
                self._refresh_folder_status()
                self._set_status("Ready", SUCCESS)
                
        threading.Thread(target=run, daemon=True).start()

    def _on_cleanse(self):
        try:
            from config import BCA_EXCEL_DIR, MANDIRI_ZIP_DIR, BRI_ZIP_DIR, ODO_EXCEL_PATH, OUTPUT_DIR
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            recap_dir = BASE_DIR / "recap" / timestamp
            
            self._log_write("\n── Cleaning Data ──\n", "head")
            
            moved_count = 0
            dirs_to_clean = [BCA_EXCEL_DIR, MANDIRI_ZIP_DIR, BRI_ZIP_DIR, OUTPUT_DIR]
            
            for d in dirs_to_clean:
                if d.exists() and d.is_dir():
                    for f in d.glob("*"):
                        if f.name == "journal_creation_log.xlsx":
                            continue
                        target_subdir = recap_dir / d.name
                        target_subdir.mkdir(parents=True, exist_ok=True)
                        shutil.move(str(f), str(target_subdir / f.name))
                        moved_count += 1
                             
            if ODO_EXCEL_PATH.exists():
                recap_dir.mkdir(parents=True, exist_ok=True)
                shutil.move(str(ODO_EXCEL_PATH), str(recap_dir / ODO_EXCEL_PATH.name))
                moved_count += 1
                
            if moved_count > 0:
                self._log_write(f"✅ {moved_count} files moved to: recap/{timestamp}/\n", "ok")
            else:
                self._log_write("ℹ️ No data to clean.\n", "dim")
                
            self._refresh_folder_status()
        except Exception as e:
            self._log_write(f"\n❌ Error during Data Cleanup: {e}\nEnsure your .env file is configured correctly!\n", "err")

    # ── Run / Scan ────────────────────────────────────────────────────────────
    def _on_scan(self):
        if self._running:
            return
        self._running = True
        selected_banks = [b.lower() for b, var in self._bank_vars.items() if var.get()]
        if not selected_banks:
            self._set_status("Select at least 1 bank!", ERROR)
            self._running = False
            return

        self._scan_btn.configure(state="disabled")
        self._run_btn.configure(state="disabled")
        self._open_btn.configure(state="disabled")
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
            self._set_status("Select at least 1 bank!", ERROR)
            self._running = False
            return

        self._scan_btn.configure(state="disabled")
        self._run_btn.configure(state="disabled", text="⏳ Processing...")
        self._open_btn.configure(state="disabled")
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")
        self._set_status("Processing...", WARN)
        self._refresh_folder_status()
        
        def run_all():
            try:
                self.after(0, self._log_write, f"\n── Starting Single Browser Auto-Recon ──\n", "head")
                
                if getattr(sys, "frozen", False):
                    cmd = [sys.executable, "--run-downloader", "--mode", "auto_recon"]
                else:
                    cmd = [_venv_python, "odoo_downloader.py", "--mode", "auto_recon"]
                    
                email = self._email_var.get().strip()
                password = self._password_var.get()
                if email and password:
                    cmd.extend(["--email", email, "--password", password])
                    
                banks_for_dl = [b for b, var in self._bank_vars.items() if var.get()]
                if banks_for_dl:
                    cmd.extend(["--banks", ",".join(banks_for_dl)])
                    
                env = os.environ.copy()
                env.pop("TCL_LIBRARY", None)
                env.pop("TK_LIBRARY", None)
                env["PYTHONIOENCODING"] = "utf-8"
                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                    encoding="utf-8",
                    errors="replace",
                    env=env
                )
                
                for line in process.stdout:
                    self.after(0, self._log_write, line)
                    
                process.wait()
                if process.returncode != 0:
                    self.after(0, self._log_write, f"\n❌ Auto-Recon Failed.\n", "err")
                    self.after(0, self._on_done, process.returncode, None)
                    return
                
                self.after(0, self._log_write, "\n✅ Auto-Recon Completed Successfully!\n", "ok")
                self.after(0, self._on_done, 0, None)
                
            except Exception as e:
                self.after(0, self._log_write, f"\n❌ Error during Auto-Recon: {e}\n", "err")
                self.after(0, self._on_done, 1, None)
                return
            
        threading.Thread(target=run_all, daemon=True).start()

    def _run_script(self, selected_banks, is_scan=False):
        if getattr(sys, "frozen", False):
            cmd = [sys.executable, "--worker"]
        elif _venv_python_path.exists():
            cmd = [_venv_python, str(BASE_DIR / "main.py")]
        else:
            cmd = [sys.executable, str(BASE_DIR / "main.py")]

        if "all" in selected_banks:
            cmd.append("--all")
            selected_banks.remove("all")
            
        if selected_banks:
            cmd.extend(["--bank"] + selected_banks)
        
        if is_scan:
            cmd.append("--scan")

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

        last_output_file = None
        if proc.stdout:
            for line in proc.stdout:
                ls = line.rstrip()
                if ls.startswith("[DATE_RANGE]|"):
                    try:
                        _, min_d, max_d = ls.split("|")
                        self.after(0, self._set_dates, min_d, max_d)
                    except Exception:
                        pass
                    continue
                    
                ls_lower = ls.lower()
                if "reconciliation_" in ls_lower and ".xlsx" in ls_lower:
                    for part in ls.split():
                        if "reconciliation_" in part.lower() and ".xlsx" in part.lower():
                            last_output_file = part.strip()
                self.after(0, self._log_write, ls + "\n", self._tag(ls))

        proc.wait()
        
        if proc.returncode == 0 and last_output_file and not is_scan:
            self.after(0, self._log_write, f"\n── Checking Journal Entries ──\n", "head")
            
            j_cmd = [sys.executable, "journal_checker.py", last_output_file] if getattr(sys, "frozen", False) else [_venv_python, "journal_checker.py", last_output_file]
            j_proc = subprocess.Popen(
                j_cmd,
                cwd=str(BASE_DIR),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=flags,
                env=env
            )
            
            if j_proc.stdout:
                for line in j_proc.stdout:
                    self.after(0, self._log_write, line)
            j_proc.wait()
            
            if j_proc.returncode != 0:
                self.after(0, self._log_write, f"\n❌ Journal Entries Check failed.\n", "err")
                
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
        self._scan_btn.configure(state="normal")
        self._run_btn.configure(state="normal", text="▶  Reconciliation")
        self._journal_btn.configure(state="normal")
        self._open_btn.configure(state="normal")
        
        if code == 0:
            self._set_status("Finished ✓", SUCCESS)
            self._last_output = output_path
            if output_path and Path(output_path).exists():
                _open_path(output_path)
        else:
            self._set_status("Failed — check logs below", ERROR)
            
    def _set_dates(self, min_d: str, max_d: str):
        try:
            from datetime import datetime
            d_from = datetime.strptime(min_d, "%Y-%m-%d").strftime("%d/%m/%Y")
            d_to = datetime.strptime(max_d, "%Y-%m-%d").strftime("%d/%m/%Y")
            if DateEntry:
                self._date_from_widget.set_date(datetime.strptime(d_from, "%d/%m/%Y"))
                self._date_to_widget.set_date(datetime.strptime(d_to, "%d/%m/%Y"))
            else:
                self._date_from_var.set(d_from)
                self._date_to_var.set(d_to)
            self._log_write(f"\n📅 [Auto-Detect] Dates updated: {d_from} - {d_to}\n", "ok")
        except Exception as e:
            self._log_write(f"\n⚠️ Failed to update dates: {e}\n", "warn")

    # ── Journal Confirmation Modal Overhaul ───────────────────────────────────
    def _on_journal(self):
        if self._running:
            return
        
        import glob
        import os
        from openpyxl import load_workbook
        
        output_files = glob.glob(str(OUTPUT_DIR / "[Rr]econciliation_*.xlsx"))
        if not output_files:
            self._set_status("No reconciliation file found", ERROR)
            return
        latest_file = max(output_files, key=os.path.getctime)
        
        top = ctk.CTkToplevel(self)
        top.title("Confirm Journal Creation")
        
        screen_width = top.winfo_screenwidth()
        screen_height = top.winfo_screenheight()
        window_width = min(1300, int(screen_width * 0.94))
        window_height = min(1000, int(screen_height * 0.90))
        center_x = max(0, int(screen_width / 2 - window_width / 2))
        center_y = max(0, int(screen_height / 2 - window_height / 2))
        top.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        top.minsize(1200, 750)
        top.configure(fg_color=BG)
        top.transient(self)
        top.grab_set()

        items = []
        journal_state = []
        
        def _load_data():
            nonlocal items, journal_state
            import glob
            import os
            output_files = glob.glob(str(OUTPUT_DIR / "[Rr]econciliation_*.xlsx"))
            if not output_files: return
            latest_file = max(output_files, key=os.path.getctime)
            try:
                wb = load_workbook(latest_file, data_only=True)
                if "Daily Summary" not in wb.sheetnames:
                    self._set_status("'Daily Summary' sheet not found", ERROR)
                    return
                def get_col_map(sheet, row_idx=3):
                    return {str(sheet.cell(row=row_idx, column=c).value).strip().lower(): c for c in range(1, sheet.max_column + 1) if sheet.cell(row=row_idx, column=c).value}

                if "Mutation Summary" in wb.sheetnames:
                    ws_mut = wb["Mutation Summary"]
                    col_map = get_col_map(ws_mut)
                    c_date = col_map.get("payment date", 3)
                    c_bank = col_map.get("bank", 4)
                    c_group = col_map.get("journal", 5)
                    c_cat = col_map.get("transaction category", 7)
                    c_amt = col_map.get("total amount", 8)
                
                    from collections import defaultdict
                    self._mutation_totals = defaultdict(float)
                    self._mutation_raw = []
                    for row in range(4, ws_mut.max_row + 1):
                        tanggal = ws_mut.cell(row=row, column=c_date).value
                        bank = ws_mut.cell(row=row, column=c_bank).value
                        group = ws_mut.cell(row=row, column=c_group).value
                        cat = ws_mut.cell(row=row, column=c_cat).value
                        amount = ws_mut.cell(row=row, column=c_amt).value
                        if tanggal and group and amount:
                            self._mutation_raw.append({"payment_date": tanggal, "bank": bank, "group": group, "category": cat, "amount": float(amount)})
                            try:
                                self._mutation_totals[(str(tanggal).strip(), str(group).strip())] += float(amount)
                            except ValueError:
                                pass
                else:
                    self._mutation_totals = {}
                
                if "Admin Fee" in wb.sheetnames:
                    ws_adm = wb["Admin Fee"]
                    col_map = get_col_map(ws_adm)
                    c_date = col_map.get("payment date", 3)
                    c_bank = col_map.get("bank", 4)
                    c_group = col_map.get("journal", 5)
                    c_cat = col_map.get("transaction category", 7)
                    c_amt = col_map.get("total amount", 8)
                
                    from collections import defaultdict
                    self._admin_totals = defaultdict(float)
                    self._admin_raw = []
                    for row in range(4, ws_adm.max_row + 1):
                        tanggal = ws_adm.cell(row=row, column=c_date).value
                        bank = ws_adm.cell(row=row, column=c_bank).value
                        group = ws_adm.cell(row=row, column=c_group).value
                        cat = ws_adm.cell(row=row, column=c_cat).value
                        amount = ws_adm.cell(row=row, column=c_amt).value
                        if tanggal and group and amount:
                            self._admin_raw.append({"payment_date": tanggal, "bank": bank, "group": group, "category": cat, "amount": float(amount)})
                            try:
                                self._admin_totals[(str(tanggal).strip(), str(group).strip())] += float(amount)
                            except ValueError:
                                pass
                else:
                    self._admin_totals = {}

                ws = wb["Daily Summary"]
                col_map = get_col_map(ws)
                c_date = col_map.get("date", 2)
                c_pdate = col_map.get("payment date", 3)
                c_bank = col_map.get("bank", 4)
                c_group = col_map.get("journal", 5)
                c_tbank = col_map.get("total bank", 6)
                c_todoo = col_map.get("total odoo", 7)
                c_diff = col_map.get("difference", 8)
                c_recon = col_map.get("reconciled", 9)
                c_status = col_map.get("status", 10)
                c_jstatus = col_map.get("journal information", 11)
            
                items.clear()
                for row in range(4, ws.max_row + 1):
                    bank = ws.cell(row=row, column=c_bank).value
                    group = ws.cell(row=row, column=c_group).value
                    tanggal = ws.cell(row=row, column=c_date).value
                    payment_date = ws.cell(row=row, column=c_pdate).value
                    total_bank = ws.cell(row=row, column=c_tbank).value
                    total_odoo = ws.cell(row=row, column=c_todoo).value
                    selisih = ws.cell(row=row, column=c_diff).value
                    reconciled = ws.cell(row=row, column=c_recon).value
                    status = ws.cell(row=row, column=c_status).value
                    journal_status = ws.cell(row=row, column=c_jstatus).value
                    if not bank or not status: continue
                
                    try:
                        diff = abs(float(selisih)) if selisih is not None else 0
                    except:
                        diff = 0
                    
                    from config import JOURNAL_TOLERANCE
                
                    status_str = str(status).strip()
                    status_valid = ("Match" in status_str) or (diff <= JOURNAL_TOLERANCE)
                
                    mutation_found = False
                    mutation_matched = False
                    mut_total = 0.0
                    try:
                        d_mut_str = str(payment_date).strip()
                    
                        mut_amount = self._mutation_totals.get((d_mut_str, str(group).strip()), 0.0)
                        adm_amount = self._admin_totals.get((d_mut_str, str(group).strip()), 0.0)
                        mut_total = mut_amount + adm_amount
                    
                        if mut_amount > 0:
                            mutation_found = True
                            mut_diff = abs(mut_total - float(total_bank)) if total_bank else 0
                            if mut_diff <= JOURNAL_TOLERANCE:
                                mutation_matched = True
                    except Exception as e:
                        pass

                    items.append({
                        "row": row,
                        "bank": bank,
                        "group": group,
                        "tanggal": tanggal,
                        "payment_date": payment_date,
                        "merchant_amount": total_bank,
                        "odoo_amount": total_odoo,
                        "amount": total_bank,
                        "selisih": selisih,
                        "mutation_found": mutation_found,
                        "mutation_matched": mutation_matched,
                        "mutation_amount": mut_total,
                        "reconciled": reconciled,
                        "journal_status": journal_status,
                        "status_valid": status_valid
                    })
                if not items:
                    self._set_status(f"No data with difference <= {JOURNAL_TOLERANCE}", ERROR)
                    return
            except Exception as e:
                self._set_status(f"Error reading excel: {e}", ERROR)
                return

            journal_state.clear()
            for item in items:
                is_reconciled = str(item.get("reconciled", "")).strip().lower() == "yes"
                status_valid = item.get("status_valid", True)
            
                disabled_edc = False
                disabled_ar = False
            
                if not is_reconciled or not status_valid:
                    disabled_edc = True
                    disabled_ar = True
            
                if not item.get("mutation_matched", False):
                    disabled_ar = True
                
                j_status = item.get("journal_status")
                if j_status:
                    j_status_str = str(j_status).strip()
                    if j_status_str not in ["", "None", "Not Yet"]:
                        parts = [p.strip() for p in j_status_str.split("|")]
                        for p in parts:
                            if "(Both" in p:
                                if "Posted" in p:
                                    disabled_edc = True
                                    disabled_ar = True
                            elif "(EDC" in p:
                                if "Posted" in p:
                                    disabled_edc = True
                            elif "(AR" in p:
                                if "Posted" in p:
                                    disabled_ar = True

                var_item = tk.BooleanVar(value=not (disabled_edc and disabled_ar))
                var_edc = tk.BooleanVar(value=not disabled_edc)
                var_ar = tk.BooleanVar(value=not disabled_ar)

                journal_state.append({
                    "item": item,
                    "var_item": var_item,
                    "var_edc": var_edc,
                    "var_ar": var_ar,
                    "disabled_edc": disabled_edc,
                    "disabled_ar": disabled_ar
                })

        _load_data()

        ITEMS_PER_PAGE = 15
        current_page = [0]
        
        # Modal Header Bar
        header_frame = ctk.CTkFrame(top, fg_color=PANEL, corner_radius=0, height=80, border_color=BORDER, border_width=1)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)
        
        left_header = ctk.CTkFrame(header_frame, fg_color="transparent")
        left_header.pack(side="left", padx=24, pady=16)
        
        ctk.CTkLabel(
            left_header, text="Confirm Journal Creation",
            font=(FONT_FAMILY, 15, "bold"), text_color=TEXT
        ).pack(anchor="w")
        ctk.CTkLabel(
            left_header, text="Review and select transactions to post. Expand any row to preview journal entries.",
            font=(FONT_FAMILY, 10), text_color=MUTED
        ).pack(anchor="w", pady=(2, 0))
        
        def _refresh_modal():
            _load_data()
            render_page(current_page[0])
            
        ctk.CTkButton(
            header_frame, text="↻ Refresh Data", height=32,
            fg_color=PANEL, hover_color=PREVIEW_BG,
            border_color=BORDER_DARK, border_width=1,
            text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
            corner_radius=6, command=_refresh_modal
        ).pack(side="right", padx=24)
        
        # Main Scrollable Body
        body_frame = ctk.CTkFrame(top, fg_color=BG, corner_radius=0)
        body_frame.pack(fill="both", expand=True, padx=24, pady=16)
        
        list_frame = ctk.CTkFrame(body_frame, fg_color=PANEL, corner_radius=8, border_color=BORDER, border_width=1)
        list_frame.pack(fill="both", expand=True)
        
        canvas = tk.Canvas(list_frame, bg=PANEL, highlightthickness=0)
        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=PANEL)
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Modal Footer Toolbar
        footer_frame = ctk.CTkFrame(top, fg_color=PANEL, corner_radius=0, height=70, border_color=BORDER, border_width=1)
        footer_frame.pack(fill="x", side="bottom")
        footer_frame.pack_propagate(False)
        
        footer_inner = ctk.CTkFrame(footer_frame, fg_color="transparent")
        footer_inner.pack(fill="both", expand=True, padx=24, pady=14)

        pagination_frame = ctk.CTkFrame(footer_inner, fg_color="transparent")
        pagination_frame.pack(side="left")
        
        btn_prev = ctk.CTkButton(
            pagination_frame, text="◄ Prev", width=70, height=32,
            fg_color=PANEL, hover_color=PREVIEW_BG,
            border_color=BORDER_DARK, border_width=1,
            text_color=TEXT, font=(FONT_FAMILY, 10, "bold"), corner_radius=6
        )
        btn_prev.pack(side="left")
        
        lbl_page = ctk.CTkLabel(pagination_frame, text="", font=(FONT_FAMILY, 10, "bold"), text_color=TEXT)
        lbl_page.pack(side="left", padx=12)
        
        btn_next = ctk.CTkButton(
            pagination_frame, text="Next ►", width=70, height=32,
            fg_color=PANEL, hover_color=PREVIEW_BG,
            border_color=BORDER_DARK, border_width=1,
            text_color=TEXT, font=(FONT_FAMILY, 10, "bold"), corner_radius=6
        )
        btn_next.pack(side="left")
        
        def _on_mousewheel(event):
            delta = event.delta
            if abs(delta) >= 120:
                delta = int(delta / 120)
            if delta != 0:
                canvas.yview_scroll(int(-1 * delta), "units")
                
        top.bind("<MouseWheel>", _on_mousewheel)
        canvas.bind("<MouseWheel>", _on_mousewheel)
        scrollable_frame.bind("<MouseWheel>", _on_mousewheel)
        
        def render_page(page_idx):
            active_det = {"btn": None, "frm": None}
            for widget in scrollable_frame.winfo_children():
                widget.destroy()
                
            # Column Minwidth Specifications for Clean Breathing Room
            col_widths = {
                0: 40,   # Expand ►
                1: 60,   # Select
                2: 110,  # Date
                3: 140,  # Journal
                4: 130,  # Merchant Amt
                5: 130,  # Odoo Amt
                6: 145,  # Mutation + Admin
                7: 80,  # Difference
                8: 54,   # EDC
                9: 115,  # EDC Status
                10: 54,  # AR
                11: 115   # AR Status
            }
            for col, w in col_widths.items():
                scrollable_frame.grid_columnconfigure(col, minsize=w)

            headers = ["", "Select", "Date", "Journal", "Merchant Amt", "Odoo Amt", "Mutation + Admin", "Difference", "EDC", "EDC Status", "AR", "AR Status"]
            for col, h in enumerate(headers):
                lbl_anchor = "w" if col in [2, 3] else "e" if col in [4, 5, 6, 7] else "center"
                lbl = tk.Label(
                    scrollable_frame, text=h, bg=PREVIEW_BG, fg=MUTED,
                    font=(FONT_FAMILY, 9, "bold"), anchor=lbl_anchor, padx=14
                )
                lbl.grid(row=0, column=col, sticky="nsew", pady=(0, 4), ipady=8)
                
            dummy = tk.Label(scrollable_frame, text="", bg=PREVIEW_BG)
            dummy.grid(row=0, column=len(headers), sticky="nsew", pady=(0, 4), ipady=8)
            scrollable_frame.grid_columnconfigure(len(headers), weight=1)
            
            # Header Bottom Border Line
            tk.Frame(scrollable_frame, bg=BORDER_DARK, height=1).grid(row=1, column=0, columnspan=len(headers)+1, sticky="ew", pady=(0, 4))
                
            start_idx = page_idx * ITEMS_PER_PAGE
            end_idx = min(start_idx + ITEMS_PER_PAGE, len(journal_state))
            
            for i, state in enumerate(journal_state[start_idx:end_idx]):
                r_main = i * 3 + 2
                r_det  = i * 3 + 3
                r_sep  = i * 3 + 4
                bg_row = PANEL if i % 2 == 0 else "#FAFAFC"
                
                item = state["item"]
                var_item = state["var_item"]
                var_edc = state["var_edc"]
                var_ar = state["var_ar"]
                
                amt = float(item['amount']) if item['amount'] else 0
                sel = float(item['selisih']) if item['selisih'] else 0
                
                def _on_jurnal_toggle(vi=var_item, ve=var_edc, va=var_ar, de=state["disabled_edc"], da=state["disabled_ar"]):
                    if not de:
                        ve.set(vi.get())
                    if not da:
                        va.set(vi.get())
                
                import config
                from collections import defaultdict
                
                b_name = str(item['bank']).lower()
                b_group = str(item['group']).lower()
                
                alias = ""
                props = {}
                for a, p in config.BANK_ACCOUNTS.get(b_name, {}).items():
                    if p.get("group", "").lower() == b_group:
                        alias = a
                        props = p
                        break
                        
                det_frame = tk.Frame(scrollable_frame, bg=PREVIEW_BG, highlightbackground=BORDER_DARK, highlightthickness=1)
                
                # EDC Section Preview
                edc_debit = props.get("edc_debit") or f"{str(item['bank']).upper()} EDC Debit"
                edc_credit = props.get("edc_credit") or f"{str(item['group'])} Credit"
                
                edc_frame = tk.Frame(det_frame, bg=PREVIEW_BG)
                edc_frame.pack(side="left", anchor="n", padx=20, pady=10)
                
                tk.Label(edc_frame, text="EDC Journal:", bg=PREVIEW_BG, fg=ACCENT, font=(FONT_FAMILY, 9, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")
                tk.Label(edc_frame, text="Debit:", bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 9)).grid(row=1, column=0, sticky="w")
                tk.Label(edc_frame, text=edc_debit, bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 9)).grid(row=1, column=1, sticky="w", padx=(10, 30))
                tk.Label(edc_frame, text=f"Rp {amt:,.0f}", bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 9, "bold")).grid(row=1, column=2, sticky="e")
                tk.Label(edc_frame, text="Credit:", bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 9)).grid(row=2, column=0, sticky="w")
                tk.Label(edc_frame, text=edc_credit, bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 9)).grid(row=2, column=1, sticky="w", padx=(10, 30))
                tk.Label(edc_frame, text=f"Rp {amt:,.0f}", bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 9, "bold")).grid(row=2, column=2, sticky="e")
                
                # AR Section Preview
                if item.get("mutation_matched", False):
                    m_date = item['payment_date']
                    m_group = item['group']
                    m_raw = [m for m in getattr(self, "_mutation_raw", []) if m["payment_date"] == m_date and m["group"] == m_group]
                    a_raw = [a for a in getattr(self, "_admin_raw", []) if a["payment_date"] == m_date and a["group"] == m_group]
                    
                    ar_debits = []
                    for m in m_raw:
                        cat = (m["category"] or "").replace(" ", "").lower()
                        acc = props.get(f"ar_debit_{cat}") or props.get("ar_debit") or f"AR DEBIT ({cat})"
                        ar_debits.append({"account": acc, "amount": m["amount"]})
                        
                    admin_sums = defaultdict(float)
                    for a in a_raw:
                        cat = (a["category"] or "").replace(" ", "").lower()
                        acc = props.get(f"admin_debit_{cat}") or props.get("admin_debit") or f"ADMIN DEBIT ({cat})"
                        admin_sums[acc] += a["amount"]
                        
                    for acc, a_amt in admin_sums.items():
                        if a_amt > 0:
                            ar_debits.append({"account": acc, "amount": a_amt})
                    
                    t_debit = sum(d["amount"] for d in ar_debits)
                    t_credit = float(amt)
                    t_diff = t_debit - t_credit
                    
                    ar_rows = []
                    for d in ar_debits:
                        ar_rows.append(("Debit:", d['account'], d['amount']))
                        
                    ar_rows.append(("Credit:", edc_debit, t_credit))
                    
                    if round(t_diff, 2) > 0:
                        ar_rows.append(("Credit:", "82005 Bank Difference Income", abs(t_diff)))
                    elif round(t_diff, 2) < 0:
                        ar_rows.append(("Debit:", "8107 Bank Difference Loss", abs(t_diff)))
                        
                    ar_frame = tk.Frame(det_frame, bg=PREVIEW_BG)
                    ar_frame.pack(side="left", anchor="n", padx=30, pady=10)
                    
                    tk.Label(ar_frame, text="AR Journal:", bg=PREVIEW_BG, fg=ACCENT, font=(FONT_FAMILY, 9, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")
                    
                    for idx, (typ, acc, amt_val) in enumerate(ar_rows):
                        r = idx + 1
                        tk.Label(ar_frame, text=typ, bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 9)).grid(row=r, column=0, sticky="w")
                        tk.Label(ar_frame, text=acc, bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 9)).grid(row=r, column=1, sticky="w", padx=(10, 30))
                        tk.Label(ar_frame, text=f"Rp {amt_val:,.0f}", bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 9, "bold")).grid(row=r, column=2, sticky="e")
                
                def _toggle_det(btn, frm=det_frame, row_idx=r_det):
                    if frm.winfo_ismapped():
                        frm.grid_remove()
                        btn.config(text="►")
                        if active_det["frm"] == frm:
                            active_det["btn"] = None
                            active_det["frm"] = None
                    else:
                        prev_frm = active_det["frm"]
                        prev_btn = active_det["btn"]
                        if prev_frm and prev_frm.winfo_exists() and prev_frm.winfo_ismapped():
                            prev_frm.grid_remove()
                            if prev_btn and prev_btn.winfo_exists():
                                prev_btn.config(text="►")
                        
                        frm.grid(row=row_idx, column=1, columnspan=10, sticky="w", pady=(8, 10))
                        btn.config(text="▼")
                        active_det["btn"] = btn
                        active_det["frm"] = frm
                        
                btn_expand = tk.Label(scrollable_frame, text="►", bg=bg_row, fg=ACCENT, cursor="hand2", font=(FONT_FAMILY, 10, "bold"))
                btn_expand.bind("<Button-1>", lambda e, b=btn_expand, f=det_frame: _toggle_det(b, f))
                btn_expand.grid(row=r_main, column=0, padx=8, ipady=6)
                
                if state["disabled_edc"] and state["disabled_ar"]:
                    cb_item = tk.Label(scrollable_frame, text="—", bg=bg_row, fg=MUTED, font=(FONT_FAMILY, 10))
                else:
                    cb_item = tk.Checkbutton(scrollable_frame, variable=var_item, bg=bg_row, selectcolor=bg_row, command=_on_jurnal_toggle)
                cb_item.grid(row=r_main, column=1, pady=3, ipady=4)
                
                tk.Label(scrollable_frame, text=str(item['tanggal']), bg=bg_row, fg=TEXT, font=(FONT_FAMILY, 9)).grid(row=r_main, column=2, sticky="w", padx=14, ipady=6)
                tk.Label(scrollable_frame, text=str(item['group']), bg=bg_row, fg=TEXT, font=(FONT_FAMILY, 9, "bold")).grid(row=r_main, column=3, sticky="w", padx=14, ipady=6)
                amt_merch = float(item.get('merchant_amount') or 0)
                amt_odoo = float(item.get('odoo_amount') or 0)
                tk.Label(scrollable_frame, text=f"Rp {amt_merch:,.0f}", bg=bg_row, fg=TEXT, font=(FONT_FAMILY, 9)).grid(row=r_main, column=4, sticky="e", padx=14, ipady=6)
                tk.Label(scrollable_frame, text=f"Rp {amt_odoo:,.0f}", bg=bg_row, fg=TEXT, font=(FONT_FAMILY, 9)).grid(row=r_main, column=5, sticky="e", padx=14, ipady=6)
                
                mut_amt = float(item.get("mutation_amount", 0))
                tk.Label(scrollable_frame, text=f"Rp {mut_amt:,.0f}", bg=bg_row, fg=TEXT, font=(FONT_FAMILY, 9)).grid(row=r_main, column=6, sticky="e", padx=14, ipady=6)
                
                sel_color = WARN if sel != 0 else TEXT
                tk.Label(scrollable_frame, text=f"Rp {sel:,.0f}", bg=bg_row, fg=sel_color, font=(FONT_FAMILY, 9, "bold" if sel != 0 else "normal")).grid(row=r_main, column=7, sticky="e", padx=14, ipady=6)
                
                if state["disabled_edc"]:
                    cb_edc = tk.Label(scrollable_frame, text="—", bg=bg_row, fg=MUTED, font=(FONT_FAMILY, 10))
                else:
                    cb_edc = tk.Checkbutton(scrollable_frame, variable=var_edc, bg=bg_row, selectcolor=bg_row)
                cb_edc.grid(row=r_main, column=8, padx=12, ipady=4)
                
                if state["disabled_ar"]:
                    cb_ar = tk.Label(scrollable_frame, text="—", bg=bg_row, fg=MUTED, font=(FONT_FAMILY, 10))
                else:
                    cb_ar = tk.Checkbutton(scrollable_frame, variable=var_ar, bg=bg_row, selectcolor=bg_row)
                cb_ar.grid(row=r_main, column=10, padx=12, ipady=4)
                
                edc_info_texts = []
                ar_info_texts = []
                is_reconciled = str(item.get("reconciled", "")).strip().lower() == "yes"
                status_valid = item.get("status_valid", True)
                
                if not is_reconciled:
                    edc_info_texts.append("⚠️ Unreconciled")
                elif not status_valid:
                    edc_info_texts.append("⚠️ Difference")
                
                if not is_reconciled:
                    ar_info_texts.append("⚠️ Unreconciled")
                elif not status_valid:
                    ar_info_texts.append("⚠️ Difference")
                elif not item.get("mutation_matched", False):
                    if not item.get("mutation_found", False):
                        ar_info_texts.append("⚠️ No Mutation")
                    else:
                        ar_info_texts.append("⚠️ Mut Difference")
                
                j_status = item.get("journal_status")
                if j_status:
                    j_status_str = str(j_status).strip()
                    if j_status_str not in ["", "None", "Not Yet"]:
                        parts = [p.strip() for p in j_status_str.split("|")]
                        for p in parts:
                            if "(Both" in p:
                                stripped = p.replace("(Both Difference)", "(Diff)").replace("(Both)", "").replace("Posted", "✅ Posted").replace("Draft", "📌 Draft").strip()
                                if not edc_info_texts:
                                    edc_info_texts.append(stripped)
                                if not ar_info_texts:
                                    ar_info_texts.append(stripped)
                            elif "(EDC" in p:
                                stripped = p.replace("(EDC Difference)", "(Diff)").replace("(EDC)", "").replace("Posted", "✅ Posted").replace("Draft", "📌 Draft").strip()
                                if not edc_info_texts:
                                    edc_info_texts.append(stripped)
                            elif "(AR" in p:
                                stripped = p.replace("(AR Difference)", "(Diff)").replace("(AR)", "").replace("Posted", "✅ Posted").replace("Draft", "📌 Draft").strip()
                                if not ar_info_texts:
                                    ar_info_texts.append(stripped)
                    
                if edc_info_texts:
                    lbl_color = SUCCESS if all(t in ["✅ Posted", "📌 Draft"] for t in edc_info_texts) else WARN
                    tk.Label(scrollable_frame, text="\n".join(edc_info_texts), bg=bg_row, fg=lbl_color, font=(FONT_FAMILY, 8, "bold")).grid(row=r_main, column=9, sticky="w", padx=14, ipady=6)
                    
                if ar_info_texts:
                    lbl_color = SUCCESS if all(t in ["✅ Posted", "📌 Draft"] for t in ar_info_texts) else WARN
                    tk.Label(scrollable_frame, text="\n".join(ar_info_texts), bg=bg_row, fg=lbl_color, font=(FONT_FAMILY, 8, "bold")).grid(row=r_main, column=11, sticky="w", padx=14, ipady=6)
                
                # Row Divider Line
                tk.Frame(scrollable_frame, bg=BORDER, height=1).grid(row=r_sep, column=0, columnspan=len(headers)+1, sticky="ew")
                
            total_pages = max(1, (len(journal_state) + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE)
            lbl_page.configure(text=f"Page {page_idx + 1} of {total_pages}")
            
            btn_prev.configure(state="normal" if page_idx > 0 else "disabled")
            btn_next.configure(state="normal" if page_idx < total_pages - 1 else "disabled")
            
            canvas.yview_moveto(0)
            
        def _prev_page():
            if current_page[0] > 0:
                current_page[0] -= 1
                render_page(current_page[0])
                
        def _next_page():
            total_pages = (len(journal_state) + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE
            if current_page[0] < total_pages - 1:
                current_page[0] += 1
                render_page(current_page[0])
                
        btn_prev.configure(command=_prev_page)
        btn_next.configure(command=_next_page)
        
        render_page(0)
        
        def _get_selected_config():
            selected = []
            for v in journal_state:
                if v["var_item"].get():
                    selected.append({
                        "row": v["item"]["row"],
                        "edc": v["var_edc"].get(),
                        "ar": v["var_ar"].get()
                    })
            return selected
            
        def _export(mode="edc"):
            selected = _get_selected_config()
            if not selected:
                messagebox.showwarning("Warning", "Select at least 1 transaction")
                return
            import json
            from journal_generator import generate_journal_import
            config_path = BASE_DIR / "journal_config.json"
            config_path.write_text(json.dumps(selected))
            try:
                out_path = generate_journal_import(latest_file, config_path, mode=mode, is_preview=True)
                if out_path:
                    messagebox.showinfo("Success", f"{mode.upper()} Journal exported:\n{out_path.name}")
                    _open_path(str(out_path))
                else:
                    messagebox.showerror("Error", f"Failed to generate {mode.upper()} journal.")
            except Exception as e:
                messagebox.showerror("Error", f"Error exporting {mode.upper()}: {e}")

        def _process():
            selected = _get_selected_config()
            if not selected:
                messagebox.showwarning("Warning", "Select at least 1 transaction")
                return
            import json
            config_path = BASE_DIR / "journal_config.json"
            config_path.write_text(json.dumps(selected))
            
            from journal_generator import generate_journal_import
            recon_file = Path(latest_file)
            if not recon_file.exists():
                messagebox.showerror("Error", f"Reconciliation file not found at {recon_file.name}. Did you rename it while this window was open?")
                return
                
            try:
                out_path = generate_journal_import(recon_file, config_path, mode="both", is_preview=True)
                if not out_path:
                    messagebox.showerror("Error", "Failed to generate combined preview file.")
                    return
            except Exception as e:
                messagebox.showerror("Error", f"Error generating combined preview: {e}")
                return
                
            ar_count = sum(1 for item in selected if item.get("ar", False))
            edc_count = sum(1 for item in selected if item.get("edc", False))
            
            def _show_custom_confirm():
                dlg = ctk.CTkToplevel(top)
                dlg.title("Confirm Upload")
                dlg.geometry("520x290")
                dlg.configure(fg_color=PANEL)
                dlg.transient(top)
                dlg.grab_set()
                
                dlg.update_idletasks()
                x = top.winfo_x() + (top.winfo_width() - 520) // 2
                y = top.winfo_y() + (top.winfo_height() - 290) // 2
                dlg.geometry(f"+{x}+{y}")
                
                ctk.CTkLabel(
                    dlg, text="Confirm Upload to Odoo",
                    font=(FONT_FAMILY, 15, "bold"), text_color=TEXT
                ).pack(pady=(24, 8))
                
                ctk.CTkLabel(
                    dlg, 
                    text=f"Ready to import to Odoo.\n\nAR Journals: {ar_count}   |   EDC Journals: {edc_count}\n\nTo make manual edits before importing, click 'Edit Excel'.", 
                    justify="center", text_color=MUTED, font=(FONT_FAMILY, 10)
                ).pack(pady=(0, 20), padx=20)
                
                result = {"confirm": False}
                
                def on_upload():
                    result["confirm"] = True
                    dlg.destroy()
                    
                def on_cancel():
                    dlg.destroy()
                
                btn_frame = ctk.CTkFrame(dlg, fg_color="transparent")
                btn_frame.pack(fill="x", padx=20, pady=10)
                
                ctk.CTkButton(
                    btn_frame, text="Cancel", height=36,
                    fg_color=PANEL, hover_color=PREVIEW_BG,
                    border_color=BORDER_DARK, border_width=1,
                    text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
                    corner_radius=6, command=on_cancel
                ).pack(side="left", padx=4, expand=True, fill="x")
                
                ctk.CTkButton(
                    btn_frame, text="Edit Excel", height=36,
                    fg_color=PANEL, hover_color=PREVIEW_BG,
                    border_color=BORDER_DARK, border_width=1,
                    text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
                    corner_radius=6, command=lambda: _open_path(str(out_path))
                ).pack(side="left", padx=4, expand=True, fill="x")
                
                ctk.CTkButton(
                    btn_frame, text="Upload to Odoo", height=36,
                    fg_color=SUCCESS, hover_color=SUCCESS_DARK,
                    text_color=WHITE, font=(FONT_FAMILY, 10, "bold"),
                    corner_radius=6, command=on_upload
                ).pack(side="left", padx=4, expand=True, fill="x")
                
                top.wait_window(dlg)
                return result["confirm"]
                
            confirm = _show_custom_confirm()
            
            if not confirm:
                return
                
            top.destroy()
            
            self._running = True
            self._set_status("Uploading Edited Journal to Odoo...", WARN)
            self._journal_btn.configure(state="disabled")
            
            def run_script():
                try:
                    env = os.environ.copy()
                    env.pop("TCL_LIBRARY", None)
                    env.pop("TK_LIBRARY", None)
                    env["PYTHONIOENCODING"] = "utf-8"
                    env["PYTHONUTF8"] = "1"
                    
                    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if IS_WINDOWS else 0
                    
                    cmd = [_venv_python, "odoo_journal_creator.py", "--file", str(recon_file), "--import-file", str(out_path), "--config", str(config_path)]
                    proc = subprocess.Popen(
                        cmd, cwd=str(BASE_DIR),
                        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                        text=True, bufsize=1, encoding="utf-8",
                        env=env, creationflags=flags
                    )
                    if proc.stdout:
                        for line in iter(proc.stdout.readline, ''):
                            self.after(0, self._log_write, line, "")
                    proc.wait()
                    if proc.returncode == 0:
                        self.after(0, self._on_done, 0, None)
                    else:
                        self.after(0, self._on_done, proc.returncode, None)
                except Exception as e:
                    self.after(0, self._log_write, f"ERROR: {str(e)}\n", "err")
                    self.after(0, self._on_done, 1, None)
            
            threading.Thread(target=run_script, daemon=True).start()
            
        footer_right = ctk.CTkFrame(footer_inner, fg_color="transparent")
        footer_right.pack(side="right")
        
        def _select_all():
            for v in journal_state:
                if not (v["disabled_edc"] and v["disabled_ar"]):
                    v["var_item"].set(True)
                if not v["disabled_edc"]:
                    v["var_edc"].set(True)
                if not v["disabled_ar"]:
                    v["var_ar"].set(True)
            render_page(current_page[0])
                
        def _deselect_all():
            for v in journal_state:
                v["var_item"].set(False)
                if not v["disabled_edc"]:
                    v["var_edc"].set(False)
                if not v["disabled_ar"]:
                    v["var_ar"].set(False)
            render_page(current_page[0])
                
        tools_frame = ctk.CTkFrame(pagination_frame, fg_color="transparent")
        tools_frame.pack(side="left", padx=(20, 0))
        
        ctk.CTkButton(
            tools_frame, text="Select All", height=32, width=80,
            fg_color=PANEL, hover_color=PREVIEW_BG,
            border_color=BORDER_DARK, border_width=1,
            text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
            corner_radius=6, command=_select_all
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            tools_frame, text="Deselect All", height=32, width=85,
            fg_color=PANEL, hover_color=PREVIEW_BG,
            border_color=BORDER_DARK, border_width=1,
            text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
            corner_radius=6, command=_deselect_all
        ).pack(side="left")
        
        ctk.CTkButton(
            footer_right, text="Cancel", height=36, width=80,
            fg_color=PANEL, hover_color=PREVIEW_BG,
            border_color=BORDER_DARK, border_width=1,
            text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
            corner_radius=6, command=top.destroy
        ).pack(side="left", padx=(0, 8))
        
        ctk.CTkButton(
            footer_right, text="Export EDC", height=36, width=95,
            fg_color=PANEL, hover_color=PREVIEW_BG,
            border_color=BORDER_DARK, border_width=1,
            text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
            corner_radius=6, command=lambda: _export("edc")
        ).pack(side="left", padx=(0, 6))
        
        ctk.CTkButton(
            footer_right, text="Export AR", height=36, width=90,
            fg_color=PANEL, hover_color=PREVIEW_BG,
            border_color=BORDER_DARK, border_width=1,
            text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
            corner_radius=6, command=lambda: _export("ar")
        ).pack(side="left", padx=(0, 12))
        
        ctk.CTkButton(
            footer_right, text="Submit", height=36, width=100,
            fg_color=SUCCESS, hover_color=SUCCESS_DARK,
            text_color=WHITE, font=(FONT_FAMILY, 10, "bold"),
            corner_radius=6, command=_process
        ).pack(side="left")

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

    def _open_mutation(self):
        from config import MUTATION_DIR
        MUTATION_DIR.mkdir(exist_ok=True)
        _open_path(str(MUTATION_DIR))


if __name__ == "__main__":
    import sys
    import io

    if len(sys.argv) > 1:
        try:
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
        except AttributeError:
            pass
            
        if sys.argv[1] == "--run-downloader":
            import odoo_downloader
            sys.argv = [sys.argv[0]] + sys.argv[2:]
            odoo_downloader.run_downloader()
            sys.exit(0)
        elif sys.argv[1] == "--run-main":
            import runpy
            runpy.run_module('main', run_name='__main__')
            sys.exit(0)
        elif sys.argv[1] == "--install-playwright":
            import runpy
            sys.argv = [sys.argv[0], "install", "chromium"]
            runpy.run_module('playwright', run_name='__main__')
            sys.exit(0)
        elif sys.argv[1] == "-m":
            import runpy
            module_name = sys.argv[2]
            sys.argv = [sys.argv[0]] + sys.argv[3:]
            runpy.run_module(module_name, run_name='__main__')
            sys.exit(0)
        elif sys.argv[1].endswith(".py"):
            import runpy
            script_name = sys.argv[1][:-3]
            sys.argv = [sys.argv[0]] + sys.argv[2:]
            runpy.run_module(script_name, run_name='__main__')
            sys.exit(0)

    app = App()
    app.mainloop()
