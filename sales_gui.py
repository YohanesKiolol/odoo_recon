from __future__ import annotations
import sys
import os
import traceback
from pathlib import Path

def _show_fatal_error(msg: str):
    try:
        log_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(".")
        (log_dir / "sales_crash.log").write_text(msg, encoding="utf-8")
    except Exception:
        pass
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Sales Portal - Fatal Error", f"Startup Error:\n\n{msg[:1500]}")
        root.destroy()
    except Exception:
        pass
    sys.exit(1)

try:
    import tkinter as tk
    from datetime import datetime, date
    import threading
    import customtkinter as ctk
    from cloud_sync import (
        is_cloud_configured,
        fetch_discrepancies,
        resolve_discrepancy,
        reopen_discrepancy,
        test_connection,
        supabase_auth_login
    )
    import odoo_inspector
except Exception:
    _show_fatal_error(traceback.format_exc())

if getattr(sys, "frozen", False):
    if sys.platform == "darwin" and ".app/Contents/MacOS" in str(sys.executable):
        BASE_DIR = Path(sys.executable).parents[2].parent
    else:
        BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).resolve().parent

# ── Design System & Color Tokens ──────────────────────────────────────────────
IS_WINDOWS = (sys.platform == "win32")
FONT_FAMILY = ("Segoe UI", "Segoe UI Emoji", "Arial") if IS_WINDOWS else "Helvetica Neue"
FONT_MONO   = "Consolas" if IS_WINDOWS else "Menlo"

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


BG          = "#F4F5F8"
PANEL       = "#FFFFFF"
HEADER_BG   = "#FAFAFC"
PREVIEW_BG  = "#F8FAFC"
BORDER      = "#E2E8F0"
BORDER_DARK = "#CBD5E1"

ACCENT      = "#6D28D9"  # Deep Violet
ACCENT_DARK = "#5B21B6"
SUCCESS     = "#059669"  # Emerald
SUCCESS_LIGHT = "#D1FAE5"
WARN        = "#D97706"  # Amber
WARN_LIGHT  = "#FEF3C7"
ERROR       = "#DC2626"
ERROR_LIGHT = "#FEE2E2"
TEXT        = "#0F172A"
MUTED       = "#64748B"
WHITE       = "#FFFFFF"

BANK_COLORS = {
    "BCA":     ("#1E40AF", "#DBEAFE"),  # Blue
    "MANDIRI": ("#92400E", "#FEF3C7"),  # Amber
    "BRI":     ("#047857", "#D1FAE5"),  # Green
    "OTHER":   ("#4B5563", "#F3F4F6"),
}

TYPE_BADGE_INFO = {
    "bank_only":         ("🏦 Bank Only", "#4338CA", "#EEF2FF"),
    "odoo_only":         ("📦 Odoo Only", "#B45309", "#FEF3C7"),
    "unreconciled_odoo": ("⚠️ Unreconciled", "#BE123C", "#FFE4E6"),
}


def _maximize_window(win):
    """Maximize a Tk/CTk window cross-platform without taskbar clipping."""
    import sys
    win.update_idletasks()
    system = sys.platform

    if system == "win32":
        try:
            import ctypes
            from ctypes import wintypes
            rect = wintypes.RECT()
            if ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0):
                work_x = rect.left
                work_y = rect.top
                work_w = rect.right - rect.left
                work_h = rect.bottom - rect.top
                client_h = max(500, work_h - 45)
                win.geometry(f"{work_w}x{client_h}+{work_x}+{work_y}")
                return
        except Exception:
            pass
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        win.geometry(f"{sw}x{max(500, sh - 80)}+0+0")
    elif system == "darwin":
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        win.geometry(f"{sw}x{max(500, sh - 95)}+0+25")
    else:
        try:
            win.attributes("-zoomed", True)
        except Exception:
            sw = win.winfo_screenwidth()
            sh = win.winfo_screenheight()
            win.geometry(f"{sw}x{max(500, sh - 80)}+0+0")


class SalesPortalApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Bank Recon — Sales Portal")
        self.geometry("1280x860")
        self.minsize(1000, 680)
        self.resizable(True, True)
        self.configure(fg_color=BG)
        self.after(50, lambda: _maximize_window(self))

        self._auth_user: dict | None = None
        self._all_items: list[dict] = []
        self._filtered_items: list[dict] = []
        self._current_type_filter = "ALL"
        self._current_bank_filter = "ALL"
        self._current_status_filter = "Pending"
        self._search_query = ""

        self._set_app_icon()

        # Container for swapping between Login and Dashboard
        self._root_container = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        self._root_container.pack(fill="both", expand=True)

        self._show_login_view()

    def _set_app_icon(self):
        """Set high-res native window icon."""
        try:
            assets = (
                Path(getattr(sys, "_MEIPASS", BASE_DIR)) / "assets"
                if getattr(sys, "frozen", False)
                else BASE_DIR / "assets"
            )
            if sys.platform.startswith("win"):
                ico = assets / "sales_app_icon.ico"
                if not ico.exists():
                    ico = assets / "app_icon.ico"
                if ico.exists():
                    self.iconbitmap(str(ico))
            else:
                png = assets / "sales_app_icon.png"
                if not png.exists():
                    png = assets / "app_icon.png"
                if png.exists():
                    from PIL import Image as _Img, ImageTk as _ImgTk
                    _icon = _ImgTk.PhotoImage(_Img.open(png).resize((256, 256)))
                    self.iconphoto(True, _icon)
                    self._icon_ref = _icon
        except Exception:
            pass

    # =========================================================================
    # ── Login View ───────────────────────────────────────────────────────────
    # =========================================================================
    def _show_login_view(self):
        for w in self._root_container.winfo_children():
            w.destroy()

        center_wrap = ctk.CTkFrame(self._root_container, fg_color="transparent")
        center_wrap.place(relx=0.5, rely=0.5, anchor="center")

        card = ctk.CTkFrame(center_wrap, fg_color=PANEL, corner_radius=12, border_color=BORDER, border_width=1, width=400)
        card.pack(padx=20, pady=20)

        ci = ctk.CTkFrame(card, fg_color="transparent")
        ci.pack(fill="both", expand=True, padx=32, pady=32)

        # Brand header
        ctk.CTkLabel(
            ci, text="Sales Portal",
            font=(FONT_FAMILY, 20, "bold"), text_color=ACCENT
        ).pack(anchor="w")
        ctk.CTkLabel(
            ci, text="Sign in to resolve unmapped bank payments",
            font=(FONT_FAMILY, 10, "bold"), text_color=MUTED
        ).pack(anchor="w", pady=(2, 16))

        # Email field
        ctk.CTkLabel(
            ci, text="Email Address",
            font=(FONT_FAMILY, 10, "bold"), text_color=TEXT
        ).pack(anchor="w")
        var_email = tk.StringVar()
        entry_email = ctk.CTkEntry(
            ci, textvariable=var_email, placeholder_text="sales@eyerizz.com",
            height=34, corner_radius=6, border_color=BORDER_DARK, fg_color=WHITE, text_color=TEXT,
            font=(FONT_FAMILY, 10, "bold")
        )
        entry_email.pack(fill="x", pady=(2, 10))

        # Password field
        ctk.CTkLabel(
            ci, text="Password",
            font=(FONT_FAMILY, 10, "bold"), text_color=TEXT
        ).pack(anchor="w")

        var_pass = tk.StringVar()
        pass_box = ctk.CTkFrame(ci, fg_color=WHITE, border_color=BORDER_DARK, border_width=1, corner_radius=6, height=34)
        pass_box.pack(fill="x", pady=(2, 12))
        pass_box.pack_propagate(False)

        entry_pass = ctk.CTkEntry(
            pass_box, textvariable=var_pass, placeholder_text="Enter password", show="•",
            height=32, border_width=0, fg_color="transparent", text_color=TEXT,
            font=(FONT_FAMILY, 10, "bold")
        )
        entry_pass.pack(side="left", fill="x", expand=True, padx=(8, 4))

        def _toggle_pass_view():
            entry_pass.configure(show="" if entry_pass.cget("show") == "•" else "•")

        btn_eye = tk.Label(pass_box, text="Show", bg=WHITE, fg=MUTED, font=(FONT_FAMILY, 8, "bold"), cursor="hand2")
        btn_eye.pack(side="right", padx=(4, 8))
        btn_eye.bind("<Button-1>", lambda e: _toggle_pass_view())

        # Error label
        lbl_err = ctk.CTkLabel(ci, text="", font=(FONT_FAMILY, 9, "bold"), text_color=ERROR, wraplength=320)
        lbl_err.pack(anchor="w", pady=(0, 8))

        # Submit button
        btn_login = ctk.CTkButton(
            ci, text="Sign In", height=38,
            fg_color=ACCENT, hover_color=ACCENT_DARK, text_color=WHITE,
            font=(FONT_FAMILY, 11, "bold"), corner_radius=6
        )
        btn_login.pack(fill="x", pady=(4, 8))

        # Demo credentials quick-fill hint (disabled for production)
        # def _fill_demo():
        #     var_email.set("sales@eyerizz.com")
        #     var_pass.set("sales123")
        #     lbl_err.configure(text="")
        #
        # lbl_demo = tk.Label(
        #     ci, text="Quick Fill Test Account (sales@eyerizz.com)",
        #     bg=PANEL, fg=ACCENT, font=(FONT_FAMILY, 8, "bold"), cursor="hand2"
        # )
        # lbl_demo.pack(anchor="center", pady=(2, 0))
        # lbl_demo.bind("<Button-1>", lambda e: _fill_demo())

        def _do_login(event=None):
            em = var_email.get().strip()
            pw = var_pass.get().strip()
            if not em or not pw:
                lbl_err.configure(text="Please enter both email and password.")
                return

            btn_login.configure(state="disabled", text="Signing In...")
            lbl_err.configure(text="")
            self.update_idletasks()

            ok, res = supabase_auth_login(em, pw)
            if ok and isinstance(res, dict):
                self._auth_user = res.get("user") or {"email": em}
                self._show_dashboard_view()
            else:
                btn_login.configure(state="normal", text="Sign In")
                err_msg = str(res) if res else "Invalid credentials"
                lbl_err.configure(text=f"{err_msg}")

        btn_login.configure(command=_do_login)
        entry_pass.bind("<Return>", _do_login)
        entry_email.bind("<Return>", lambda e: entry_pass.focus())
        entry_email.focus()

    # =========================================================================
    # ── Main Dashboard View ──────────────────────────────────────────────────
    # =========================================================================
    def _show_dashboard_view(self):
        for w in self._root_container.winfo_children():
            w.destroy()

        # ── Top Navigation Bar ──
        topbar = ctk.CTkFrame(self._root_container, fg_color=HEADER_BG, corner_radius=0, height=64, border_color=BORDER, border_width=1)
        topbar.pack(fill="x")
        topbar.pack_propagate(False)

        top_inner = ctk.CTkFrame(topbar, fg_color="transparent")
        top_inner.pack(fill="both", expand=True, padx=20, pady=10)

        # Left branding
        brand_f = ctk.CTkFrame(top_inner, fg_color="transparent")
        brand_f.pack(side="left", fill="y")
        ctk.CTkLabel(
            brand_f, text="Sales Discrepancy Portal",
            font=(FONT_FAMILY, 16, "bold"), text_color=ACCENT
        ).pack(anchor="w")

        user_email = (self._auth_user.get("email") if self._auth_user else "sales") or "Sales"
        user_name = self._auth_user.get("user_metadata", {}).get("name") or user_email.split("@")[0].capitalize()
        ctk.CTkLabel(
            brand_f, text=f"Logged in as {user_name} ({user_email})",
            font=(FONT_FAMILY, 10, "bold"), text_color=MUTED
        ).pack(anchor="w")

        # Right status & actions
        right_top = ctk.CTkFrame(top_inner, fg_color="transparent")
        right_top.pack(side="right", fill="y")

        self._cloud_badge = ctk.CTkLabel(
            right_top, text="● Cloud Connected",
            font=(FONT_FAMILY, 10, "bold"), text_color=SUCCESS
        )
        self._cloud_badge.pack(side="left", padx=(0, 12))

        btn_refresh = ctk.CTkButton(
            right_top, text="Refresh", height=30, width=75,
            fg_color=WHITE, hover_color=PREVIEW_BG, border_color=BORDER_DARK, border_width=1,
            text_color=TEXT, font=(FONT_FAMILY, 9, "bold"), corner_radius=5,
            command=self._load_data
        )
        btn_refresh.pack(side="left", padx=(0, 8))

        btn_logout = ctk.CTkButton(
            right_top, text="Logout", height=30, width=70,
            fg_color=WHITE, hover_color=ERROR_LIGHT, border_color=BORDER_DARK, border_width=1,
            text_color=ERROR, font=(FONT_FAMILY, 9, "bold"), corner_radius=5,
            command=self._show_login_view
        )
        btn_logout.pack(side="left")

        # ── Filter & Quick Search Bar ──
        filter_bar = ctk.CTkFrame(self._root_container, fg_color=PANEL, corner_radius=0, height=62, border_color=BORDER, border_width=1)
        filter_bar.pack(fill="x")
        filter_bar.pack_propagate(False)

        f_inner = ctk.CTkFrame(filter_bar, fg_color="transparent")
        f_inner.pack(fill="both", expand=True, padx=16, pady=10)

        # Left controls: Type tabs + Bank pills
        f_left = ctk.CTkFrame(f_inner, fg_color="transparent")
        f_left.pack(side="left", fill="y")

        self._type_btns = {}
        for t_key, t_lbl in [
            ("ALL", "All Types"),
            ("bank_only", "🏦 Bank Only"),
            ("odoo_only", "📦 Odoo Only"),
            ("unreconciled_odoo", "⚠️ Unreconciled"),
        ]:
            is_sel = (t_key == self._current_type_filter)
            t_btn = ctk.CTkButton(
                f_left, text=t_lbl, height=36, width=135 if t_key == "unreconciled_odoo" else (125 if t_key == "bank_only" else (120 if t_key == "odoo_only" else 95)),
                fg_color=ACCENT if is_sel else WHITE,
                text_color=WHITE if is_sel else TEXT,
                border_color=ACCENT if is_sel else BORDER_DARK,
                border_width=1,
                hover_color=ACCENT_DARK if is_sel else PREVIEW_BG,
                font=(FONT_FAMILY, 11, "bold"), corner_radius=6,
                command=lambda k=t_key: self._set_type_filter(k)
            )
            t_btn.pack(side="left", padx=(0, 6))
            self._type_btns[t_key] = t_btn

        # Subtle vertical separator
        sep = tk.Frame(f_left, bg=BORDER_DARK, width=1, height=24)
        sep.pack(side="left", padx=8, pady=4)

        # Bank Selector Pills
        self._bank_btns = {}
        for b_name in ["ALL", "BCA", "MANDIRI", "BRI"]:
            is_b_sel = (b_name == self._current_bank_filter)
            b_btn = ctk.CTkButton(
                f_left, text=b_name, height=36, width=55 if b_name != "MANDIRI" else 80,
                fg_color=ACCENT if is_b_sel else WHITE,
                text_color=WHITE if is_b_sel else TEXT,
                border_color=ACCENT if is_b_sel else BORDER_DARK,
                border_width=1, hover_color=ACCENT_DARK if is_b_sel else PREVIEW_BG,
                font=(FONT_FAMILY, 11, "bold"), corner_radius=6,
                command=lambda bn=b_name: self._set_bank_filter(bn)
            )
            b_btn.pack(side="left", padx=(0, 4))
            self._bank_btns[b_name] = b_btn

        # Right controls: Status tabs + Search box
        f_right = ctk.CTkFrame(f_inner, fg_color="transparent")
        f_right.pack(side="right", fill="y")

        # Search Box
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *args: self._on_search())
        search_entry = ctk.CTkEntry(
            f_right, textvariable=self._search_var, placeholder_text="🔍 Search...",
            height=36, width=180, corner_radius=6, border_color=BORDER_DARK, fg_color=WHITE, text_color=TEXT,
            font=(FONT_FAMILY, 11, "bold")
        )
        search_entry.pack(side="right", padx=(8, 0))

        # Status Tabs
        self._status_btns = {}
        for key, label in [("Pending", "Action Needed"), ("Resolve", "Resolved"), ("ALL", "All")]:
            is_st_sel = (key == self._current_status_filter)
            btn = ctk.CTkButton(
                f_right, text=label, height=36, width=115 if key == "Pending" else (90 if key == "Resolve" else 60),
                fg_color=ACCENT if is_st_sel else WHITE,
                text_color=WHITE if is_st_sel else TEXT,
                border_color=ACCENT if is_st_sel else BORDER_DARK,
                border_width=1,
                hover_color=ACCENT_DARK if is_st_sel else PREVIEW_BG,
                font=(FONT_FAMILY, 11, "bold"), corner_radius=6,
                command=lambda k=key: self._set_status_filter(k)
            )
            btn.pack(side="left", padx=(0, 4))
            self._status_btns[key] = btn

        # ── Fast Scrollable Cards Canvas ──
        self._scroll_list = ctk.CTkScrollableFrame(
            self._root_container, fg_color="transparent",
            scrollbar_button_color=BORDER, scrollbar_button_hover_color=MUTED
        )
        self._scroll_list.pack(fill="both", expand=True, padx=20, pady=8)

        # ── Compact Bottom Summary Footer ──
        footer = ctk.CTkFrame(self._root_container, fg_color=PANEL, corner_radius=0, height=38, border_color=BORDER, border_width=1)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        foot_inner = ctk.CTkFrame(footer, fg_color="transparent")
        foot_inner.pack(fill="both", expand=True, padx=20, pady=6)

        self._foot_count_lbl = ctk.CTkLabel(
            foot_inner, text="Showing 0 items",
            font=(FONT_FAMILY, 9, "bold"), text_color=MUTED
        )
        self._foot_count_lbl.pack(side="left")

        self._foot_amount_lbl = ctk.CTkLabel(
            foot_inner, text="Total: Rp 0",
            font=(FONT_FAMILY, 10, "bold"), text_color=TEXT
        )
        self._foot_amount_lbl.pack(side="right")

        self._load_data()

    def _set_type_filter(self, key):
        self._current_type_filter = key
        for k, btn in self._type_btns.items():
            if k == key:
                btn.configure(fg_color=ACCENT, text_color=WHITE, border_color=ACCENT, hover_color=ACCENT_DARK)
            else:
                btn.configure(fg_color=WHITE, text_color=TEXT, border_color=BORDER_DARK, hover_color=PREVIEW_BG)
        self._apply_filter()

    def _set_status_filter(self, key):
        self._current_status_filter = key
        for k, btn in self._status_btns.items():
            if k == key:
                btn.configure(fg_color=ACCENT, text_color=WHITE, border_color=ACCENT, hover_color=ACCENT_DARK)
            else:
                btn.configure(fg_color=WHITE, text_color=TEXT, border_color=BORDER_DARK, hover_color=PREVIEW_BG)
        self._apply_filter()

    def _set_bank_filter(self, bank):
        self._current_bank_filter = bank
        for b, btn in self._bank_btns.items():
            if b == bank:
                btn.configure(fg_color=ACCENT, text_color=WHITE, border_color=ACCENT, hover_color=ACCENT_DARK)
            else:
                btn.configure(fg_color=WHITE, text_color=TEXT, border_color=BORDER_DARK, hover_color=PREVIEW_BG)
        self._apply_filter()

    def _on_search(self):
        self._search_query = self._search_var.get().strip().lower()
        self._apply_filter()

    def _load_data(self):
        if not is_cloud_configured():
            self._cloud_badge.configure(text="● Cloud Not Configured", text_color=ERROR)
            self._render_empty_state("⚠️ Supabase is not configured in .env.")
            return

        self._cloud_badge.configure(text="● Loading...", text_color=WARN)
        self._render_empty_state("⏳ Loading discrepancy records from cloud...")

        import threading
        def _bg_fetch():
            items = fetch_discrepancies(limit=300)
            def _apply():
                self._all_items = items
                self._cloud_badge.configure(text="● Cloud Connected", text_color=SUCCESS)
                self._apply_filter()
            if self.winfo_exists():
                self.after(0, _apply)

        threading.Thread(target=_bg_fetch, daemon=True).start()

    def _apply_filter(self):
        filtered = []
        for item in self._all_items:
            dtype = str(item.get("discrepancy_type") or "bank_only").strip()
            if self._current_type_filter != "ALL" and dtype != self._current_type_filter:
                continue

            st_raw = str(item.get("status") or "Pending").strip()
            st = "Pending" if "pending" in st_raw.lower() else ("Resolve" if "resolve" in st_raw.lower() else ("Ignored" if "ignore" in st_raw.lower() else st_raw))
            if self._current_status_filter != "ALL" and st != self._current_status_filter:
                continue

            b = str(item.get("bank_name", "")).upper()
            if self._current_bank_filter != "ALL" and b != self._current_bank_filter:
                continue

            if self._search_query:
                amt_str = f"{float(item.get('amount', 0)):,.0f}".replace(",", "")
                searchable = f"{item.get('bank_name', '')} {item.get('journal', '')} {item.get('bank_number', '')} {item.get('odoo_number', '')} {item.get('sales_notes', '')} {item.get('odoo_reference', '')} {dtype} {amt_str}".lower()
                if self._search_query not in searchable:
                    continue

            filtered.append(item)

        def _sales_sort_key(it):
            b = str(it.get("bank_name", it.get("bank", ""))).strip().upper()
            d_raw = str(it.get("transaction_date", it.get("date", ""))).strip()
            if "/" in d_raw:
                pts = d_raw.split("/")
                if len(pts) == 3:
                    d_norm = f"{pts[2]}{pts[1].zfill(2)}{pts[0].zfill(2)}"
                else:
                    d_norm = d_raw
            else:
                d_norm = d_raw.replace("-", "")
            return (b, d_norm)

        filtered.sort(key=_sales_sort_key)

        self._display_limit = 100
        self._filtered_items = filtered
        self._render_cards()

    # =========================================================================
    # ── High-Performance Compact Card Renderer ──────────────────────────────
    # =========================================================================
    def _render_cards(self):
        for w in self._scroll_list.winfo_children():
            w.destroy()

        if not self._filtered_items:
            self._render_empty_state("✨ No discrepancy items match current filters.")
            self._foot_count_lbl.configure(text="Showing 0 items")
            self._foot_amount_lbl.configure(text="Total: Rp 0")
            return

        total_sum = sum(float(item.get("amount", 0.0)) for item in self._filtered_items)
        slice_items = self._filtered_items[:self._display_limit]

        # Use native lightweight tk.Frame + tk.Label for instant 0ms lag-free rendering
        for item in slice_items:
            self._build_compact_card(item)

        if len(self._filtered_items) > len(slice_items):
            remaining = len(self._filtered_items) - len(slice_items)
            more_frame = tk.Frame(self._scroll_list, bg=BG)
            more_frame.pack(fill="x", pady=10)
            def _show_more():
                self._display_limit += 100
                self._render_cards()
            btn_more = ctk.CTkButton(
                more_frame, text=f"⬇️ Load Next 100 Items (Showing {len(slice_items)} of {len(self._filtered_items)} items — {remaining} remaining)",
                height=32, fg_color=WHITE, hover_color=PREVIEW_BG, border_color=BORDER_DARK, border_width=1,
                text_color=TEXT, font=(FONT_FAMILY, 9, "bold"), command=_show_more
            )
            btn_more.pack(pady=4)

        count_str = f"Showing {len(slice_items)} of {len(self._filtered_items)} item(s)"
        self._foot_count_lbl.configure(text=count_str)
        self._foot_amount_lbl.configure(text=f"Total: Rp {total_sum:,.2f}")

        # Reset scroll position to top
        def _reset_scroll():
            try:
                self._scroll_list._parent_canvas.yview_moveto(0.0)
            except Exception:
                pass
        _reset_scroll()
        self.after(10, _reset_scroll)
        self.after(50, _reset_scroll)

    def _build_compact_card(self, item: dict):
        card = tk.Frame(self._scroll_list, bg=PANEL, highlightbackground=BORDER, highlightthickness=1, bd=0)
        card.pack(fill="x", padx=4, pady=3)
        
        # Enforce exact column positions across all rows
        card.columnconfigure(0, minsize=120)  # Col 0: Discrepancy Type Badge
        card.columnconfigure(1, minsize=80)   # Col 1: Bank Badge
        card.columnconfigure(2, minsize=95)   # Col 2: Date
        card.columnconfigure(3, minsize=140)  # Col 3: Journal
        card.columnconfigure(4, weight=1)     # Col 4: Ref / Doc (Expands & absorbs available space)
        card.columnconfigure(5, minsize=155)  # Col 5: Amount
        card.columnconfigure(6, minsize=115)  # Col 6: Action Button

        dtype = str(item.get("discrepancy_type") or "bank_only").strip()
        t_lbl, t_fg, t_bg = TYPE_BADGE_INFO.get(dtype, ("Discrepancy", TEXT, BG))
        lbl_tbadge = tk.Label(card, text=f" {t_lbl} ", bg=t_bg, fg=t_fg, font=(FONT_FAMILY, 9, "bold"), padx=5, pady=3, width=14, anchor="center")
        lbl_tbadge.grid(row=0, column=0, padx=(10, 6), pady=7, sticky="w")

        bank_name = str(item.get("bank_name", "OTHER")).upper()
        fg_col, bg_col = BANK_COLORS.get(bank_name, BANK_COLORS["OTHER"])
        lbl_badge = tk.Label(card, text=f" {bank_name} ", bg=bg_col, fg=fg_col, font=(FONT_FAMILY, 9, "bold"), padx=5, pady=3, width=7, anchor="center")
        lbl_badge.grid(row=0, column=1, padx=(0, 6), pady=7, sticky="w")

        t_date = str(item.get("transaction_date", ""))
        lbl_date = tk.Label(card, text=t_date, bg=PANEL, fg=TEXT, font=(FONT_FAMILY, 10, "bold"), width=10, anchor="w")
        lbl_date.grid(row=0, column=2, padx=(0, 6), pady=7, sticky="w")

        journal = str(item.get("journal", "General Bank"))
        lbl_j = tk.Label(card, text=journal, bg=PANEL, fg=MUTED, font=(FONT_FAMILY, 9), width=16, anchor="w")
        lbl_j.grid(row=0, column=3, padx=(0, 6), pady=7, sticky="w")

        # Ref / Doc based on category (Unclipped, expands into available space)
        if dtype == "bank_only":
            bank_ref = item.get("bank_number") or "-"
            ref_str = f"Ref: {bank_ref}"
        elif dtype == "unreconciled_odoo":
            o_ref = item.get("odoo_number") or item.get("odoo_reference") or "-"
            ref_str = f"Unrecon: {o_ref}"
        else:
            o_ref = item.get("odoo_number") or item.get("odoo_reference") or "-"
            ref_str = f"Doc: {o_ref}"

        lbl_ref = tk.Label(card, text=ref_str, bg=PANEL, fg=TEXT, font=(FONT_MONO, 10, "bold"), anchor="w")
        lbl_ref.grid(row=0, column=4, padx=(0, 10), pady=7, sticky="ew")

        amt = float(item.get("amount", 0.0))

        # Col 5: Amount
        lbl_amt = tk.Label(card, text=f"Rp {amt:,.2f}", bg=PANEL, fg=TEXT, font=(FONT_FAMILY, 11, "bold"), width=16, anchor="e")
        lbl_amt.grid(row=0, column=5, padx=(0, 10), pady=7, sticky="e")

        # Col 6: Action Button (Live Odoo Inspection)
        btn_act = ctk.CTkButton(
            card, text="🔍 Check Odoo", height=28, width=104,
            fg_color=WHITE, hover_color=PREVIEW_BG, border_color=BORDER_DARK, border_width=1,
            text_color=TEXT, font=(FONT_FAMILY, 9, "bold"), corner_radius=6,
            command=lambda it=item: self._open_diagnostic_modal(it)
        )
        btn_act.grid(row=0, column=6, padx=(0, 10), pady=7, sticky="e")

        # Fast hover effect on entire row & double-click inspection
        labels_to_hover = [lbl_date, lbl_j, lbl_ref, lbl_amt]

        def _on_enter(e):
            card.config(bg=PREVIEW_BG, highlightbackground=ACCENT)
            for l in labels_to_hover:
                l.config(bg=PREVIEW_BG)

        def _on_leave(e):
            card.config(bg=PANEL, highlightbackground=BORDER)
            for l in labels_to_hover:
                l.config(bg=PANEL)

        card.bind("<Enter>", _on_enter)
        card.bind("<Leave>", _on_leave)
        card.bind("<Double-1>", lambda e, it=item: self._open_diagnostic_modal(it))
        for l in labels_to_hover:
            l.bind("<Double-1>", lambda e, it=item: self._open_diagnostic_modal(it))

    def _render_empty_state(self, message: str):
        wrap = tk.Frame(self._scroll_list, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        wrap.pack(fill="x", padx=16, pady=24)
        tk.Label(wrap, text=message, bg=PANEL, fg=MUTED, font=(FONT_FAMILY, 10, "bold")).pack(pady=20)

    # =========================================================================
    # ── Live Odoo Diagnostics Modal ──────────────────────────────────────────
    # =========================================================================
    def _open_diagnostic_modal(self, item: dict):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Odoo Discrepancy Diagnostics")
        dlg.geometry("820x460")
        dlg.minsize(760, 400)
        dlg.resizable(True, True)
        dlg.transient(self)

        sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
        cx, cy = max(0, int(sw / 2 - 820 / 2)), max(0, int(sh / 2 - 460 / 2))
        dlg.geometry(f"820x460+{cx}+{cy}")
        dlg.configure(fg_color=BG)

        def _safe_grab():
            try:
                if dlg.winfo_exists():
                    dlg.grab_set()
            except Exception:
                pass
        dlg.after(100, _safe_grab)

        # Scrollable root container for responsiveness
        scroll_root = ctk.CTkScrollableFrame(dlg, fg_color="transparent")
        scroll_root.pack(fill="both", expand=True, padx=16, pady=16)

        content = ctk.CTkFrame(scroll_root, fg_color=PANEL, corner_radius=10, border_color=BORDER, border_width=1)
        content.pack(fill="both", expand=True)

        # ── 1. Header Info ──
        hdr = ctk.CTkFrame(content, fg_color=PREVIEW_BG, corner_radius=8, border_color=BORDER, border_width=1)
        hdr.pack(fill="x", padx=16, pady=(16, 12))

        amt = float(item.get("amount", 0.0))
        t_date = item.get("transaction_date", "")
        b_name = item.get("bank_name", "")
        journal = item.get("journal", "")
        dtype = str(item.get("discrepancy_type") or "bank_only").strip()
        t_lbl, t_fg, t_bg = TYPE_BADGE_INFO.get(dtype, ("Discrepancy", TEXT, BG))

        ctk.CTkLabel(hdr, text=f"Rp {amt:,.2f}  •  {t_lbl}", font=(FONT_FAMILY, 16, "bold"), text_color=TEXT).pack(anchor="w", padx=14, pady=(10, 3))
        ctk.CTkLabel(hdr, text=f"{b_name} • {journal} • Date: {t_date}", font=(FONT_FAMILY, 11, "bold"), text_color=ACCENT).pack(anchor="w", padx=14)

        ref_txt = f"Bank Ref: {item.get('bank_number') or '-'}"
        if item.get("odoo_number"):
            ref_txt += f"  •  Odoo Doc: {item.get('odoo_number')}"
        ctk.CTkLabel(hdr, text=ref_txt, font=(FONT_MONO, 11), text_color=MUTED).pack(anchor="w", padx=14, pady=(3, 10))

        # ── 2. Odoo Live Diagnostics Assistant ──
        diag_frame = ctk.CTkFrame(content, fg_color="#F8FAFC", corner_radius=8, border_color=BORDER_DARK, border_width=1)
        diag_frame.pack(fill="x", padx=16, pady=(0, 16))

        diag_hdr = ctk.CTkFrame(diag_frame, fg_color="transparent")
        diag_hdr.pack(fill="x", padx=14, pady=(10, 6))

        ctk.CTkLabel(diag_hdr, text="🔍 Odoo Live Diagnostics (Read-Only)", font=(FONT_FAMILY, 12, "bold"), text_color=ACCENT_DARK).pack(side="left")
        lbl_status = ctk.CTkLabel(diag_hdr, text="⏳ Checking Odoo...", font=(FONT_FAMILY, 11, "bold"), text_color=MUTED)
        lbl_status.pack(side="right")

        diag_body = ctk.CTkFrame(diag_frame, fg_color="transparent")
        diag_body.pack(fill="x", padx=14, pady=(0, 14))

        btn_bar = ctk.CTkFrame(content, fg_color="transparent")
        btn_bar.pack(fill="x", side="bottom", padx=16, pady=(0, 16))

        btn_close = ctk.CTkButton(btn_bar, text="Close", height=32, width=90, fg_color=WHITE, hover_color=PREVIEW_BG, border_color=BORDER_DARK, border_width=1, text_color=TEXT, font=(FONT_FAMILY, 10, "bold"), command=dlg.destroy)
        btn_close.pack(side="right")

        # ── 3. Background Odoo Diagnostic Inspector Worker ──
        def _apply_diag_ui(res: dict):
            if not dlg.winfo_exists():
                return
            for w in diag_body.winfo_children():
                w.destroy()

            if not res.get("success"):
                lbl_status.configure(text="⚠️ Odoo Check Offline", text_color=WARN)
                err_msg = res.get("error") or "Could not connect to Odoo server."
                ctk.CTkLabel(diag_body, text=err_msg, font=(FONT_FAMILY, 11), text_color=MUTED, wraplength=640, justify="left").pack(anchor="w", pady=4)
                return

            lbl_status.configure(text="✅ Checked Live", text_color=SUCCESS)

            # Bank Only UI
            if dtype == "bank_only":
                invoices = res.get("invoices", [])
                summary = res.get("summary", "")
                ctk.CTkLabel(diag_body, text=summary, font=(FONT_FAMILY, 12, "bold"), text_color=TEXT, wraplength=640, justify="left").pack(anchor="w", pady=(0, 8))

                if invoices:
                    for inv in invoices:
                        inv_card = ctk.CTkFrame(diag_body, fg_color=WHITE, corner_radius=8, border_color=BORDER, border_width=1)
                        inv_card.pack(fill="x", pady=4)

                        top_row = ctk.CTkFrame(inv_card, fg_color="transparent")
                        top_row.pack(fill="x", padx=12, pady=(10, 4))

                        badge_color = SUCCESS if inv["status_code"] == "AVAILABLE_OPEN" else (WARN if inv["status_code"] == "DRAFT_INVOICE" else ERROR)
                        ctk.CTkLabel(top_row, text=f" {inv['badge']} ", font=(FONT_FAMILY, 10, "bold"), text_color=WHITE, fg_color=badge_color, corner_radius=4).pack(side="left", padx=(0, 8))
                        ctk.CTkLabel(top_row, text=inv['name'], font=(FONT_MONO, 13, "bold"), text_color=TEXT).pack(side="left")
                        ctk.CTkLabel(top_row, text=f"Rp {inv['amount']:,.2f}", font=(FONT_FAMILY, 13, "bold"), text_color=SUCCESS_DARK).pack(side="right")

                        desc = f"Customer: {inv['customer']}  •  Date: {inv['date']}"
                        ctk.CTkLabel(inv_card, text=desc, font=(FONT_FAMILY, 11, "bold"), text_color=MUTED).pack(anchor="w", padx=12, pady=(0, 4))
                        if inv.get("detail"):
                            ctk.CTkLabel(inv_card, text=f"↳ {inv['detail']}", font=(FONT_FAMILY, 11), text_color=TEXT, wraplength=620, justify="left").pack(anchor="w", padx=12, pady=(0, 10))

            # Odoo Only UI
            elif dtype == "odoo_only":
                summary = res.get("summary", "")
                linked_invoices = res.get("linked_invoices", [])
                ctk.CTkLabel(diag_body, text=summary, font=(FONT_FAMILY, 12, "bold"), text_color=TEXT, wraplength=640, justify="left").pack(anchor="w", pady=(0, 8))

                if linked_invoices:
                    for li in linked_invoices:
                        li_box = ctk.CTkFrame(diag_body, fg_color=WHITE, corner_radius=8, border_color=BORDER, border_width=1)
                        li_box.pack(fill="x", pady=4)

                        r1 = ctk.CTkFrame(li_box, fg_color="transparent")
                        r1.pack(fill="x", padx=12, pady=(10, 4))
                        ctk.CTkLabel(r1, text="🔗 Linked Invoice:", font=(FONT_FAMILY, 11, "bold"), text_color=ACCENT).pack(side="left", padx=(0, 6))
                        ctk.CTkLabel(r1, text=li['name'], font=(FONT_MONO, 13, "bold"), text_color=TEXT).pack(side="left")
                        ctk.CTkLabel(r1, text=f"Rp {li['amount']:,.2f}", font=(FONT_FAMILY, 13, "bold"), text_color=TEXT).pack(side="right")

                        desc = f"Customer: {li['customer']}  •  Invoice Date: {li['date']}  •  State: {li['state']}"
                        ctk.CTkLabel(li_box, text=desc, font=(FONT_FAMILY, 11, "bold"), text_color=MUTED).pack(anchor="w", padx=12, pady=(0, 10))
                else:
                    no_inv_box = ctk.CTkFrame(diag_body, fg_color=WHITE, corner_radius=6, border_color=BORDER, border_width=1)
                    no_inv_box.pack(fill="x", pady=3)
                    ctk.CTkLabel(no_inv_box, text="⚪ No linked invoice found on this payment document in Odoo.", font=(FONT_FAMILY, 11, "bold"), text_color=MUTED).pack(anchor="w", padx=12, pady=10)

            # Unreconciled UI
            elif dtype == "unreconciled_odoo":
                summary = res.get("summary", "")
                invoices = res.get("invoices", [])
                ctk.CTkLabel(diag_body, text=summary, font=(FONT_FAMILY, 12, "bold"), text_color=TEXT, wraplength=640, justify="left").pack(anchor="w", pady=(0, 8))

                if invoices:
                    for inv in invoices:
                        inv_card = ctk.CTkFrame(diag_body, fg_color=WHITE, corner_radius=8, border_color=BORDER, border_width=1)
                        inv_card.pack(fill="x", pady=4)

                        r1 = ctk.CTkFrame(inv_card, fg_color="transparent")
                        r1.pack(fill="x", padx=12, pady=(10, 4))
                        ctk.CTkLabel(r1, text=f" {inv['badge']} ", font=(FONT_FAMILY, 10, "bold"), text_color=WHITE, fg_color=ERROR if inv.get("other_payments") else SUCCESS, corner_radius=4).pack(side="left", padx=(0, 8))
                        ctk.CTkLabel(r1, text=f"Invoice: {inv['name']}", font=(FONT_MONO, 13, "bold"), text_color=TEXT).pack(side="left")
                        ctk.CTkLabel(r1, text=f"Rp {inv['amount']:,.2f}", font=(FONT_FAMILY, 13, "bold"), text_color=TEXT).pack(side="right")

                        desc = f"Customer: {inv['customer']}  •  Date: {inv['date']}"
                        ctk.CTkLabel(inv_card, text=desc, font=(FONT_FAMILY, 11, "bold"), text_color=MUTED).pack(anchor="w", padx=12, pady=(0, 4))

                        if inv.get("other_payments"):
                            other_box = ctk.CTkFrame(inv_card, fg_color=ERROR_LIGHT, corner_radius=6, border_color=ERROR, border_width=1)
                            other_box.pack(fill="x", padx=12, pady=(4, 10))
                            ctk.CTkLabel(other_box, text="⚠️ Already linked to other payment(s):", font=(FONT_FAMILY, 11, "bold"), text_color=ERROR).pack(anchor="w", padx=8, pady=(6, 2))
                            for op in inv["other_payments"]:
                                op_txt = f"• {op['ref']}  |  Date: {op['date']}  |  Journal: {op['journal']}  |  Rp {op['amount']:,.2f}"
                                ctk.CTkLabel(other_box, text=op_txt, font=(FONT_FAMILY, 11, "bold"), text_color=TEXT, justify="left").pack(anchor="w", padx=12, pady=1)
                        else:
                            ctk.CTkLabel(inv_card, text="🟢 Invoice is not linked to any other customer payment.", font=(FONT_FAMILY, 11, "bold"), text_color=SUCCESS).pack(anchor="w", padx=12, pady=(2, 10))

            else:
                summary = res.get("summary", "")
                ctk.CTkLabel(diag_body, text=summary, font=(FONT_FAMILY, 9), text_color=MUTED).pack(anchor="w")

        def _bg_inspect():
            try:
                res = odoo_inspector.inspect_discrepancy(item)
            except Exception as e:
                res = {"success": False, "error": str(e)}
            if dlg.winfo_exists():
                dlg.after(0, lambda: _apply_diag_ui(res))

        threading.Thread(target=_bg_inspect, daemon=True).start()


if __name__ == "__main__":
    try:
        app = SalesPortalApp()
        app.mainloop()
    except Exception as e:
        import traceback
        err_str = traceback.format_exc()
        try:
            log_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(".")
            (log_dir / "sales_crash.log").write_text(err_str, encoding="utf-8")
        except Exception:
            pass
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Sales Portal - Error", f"Startup Error:\n\n{err_str[:1200]}")
            root.destroy()
        except Exception:
            pass
        sys.exit(1)

