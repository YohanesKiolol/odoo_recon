"""
Cloud Database Tab — Executive Analytics & Financial Insights Dashboard.
Presents high-level business perspective, outlet distribution, bank channel mix, and daily sales velocity.
"""
import threading
import calendar
from datetime import datetime, timedelta, date
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from PIL import Image, ImageDraw, ImageTk



from ui.theme import (
    PANEL, PREVIEW_BG, BORDER, BORDER_DARK, ACCENT, ACCENT_DARK, SUCCESS, SUCCESS_DARK,
    ERROR, WARN, TEXT, MUTED, WHITE, FONT_FAMILY, FONT_MONO, IS_WINDOWS
)
from ui.widgets.date_picker import CTkDateInput



class CloudDashboardView(ctk.CTkFrame):
    def __init__(self, master, on_sync=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.on_sync = on_sync
        self._chart_photo = None
        self._build_ui()





    def _build_ui(self):
        # 0. Global Filter Toolbar (Aligned with 16px Grid Edge)
        top_filter_bar = ctk.CTkFrame(self, fg_color=PANEL, corner_radius=8, border_color=BORDER_DARK, border_width=1, height=44)
        top_filter_bar.pack(fill="x", padx=16, pady=(6, 6))
        top_filter_bar.pack_propagate(False)

        # Left side: Filter Scope
        f_left = ctk.CTkFrame(top_filter_bar, fg_color="transparent")
        f_left.pack(side="left", padx=(10, 0))

        ico_font = ("Segoe UI Emoji", 11) if IS_WINDOWS else (FONT_FAMILY, 11)
        ctk.CTkLabel(f_left, text="🎯 Dashboard Scope:", font=(FONT_FAMILY, 10, "bold"), text_color=MUTED, fg_color="transparent").pack(side="left", padx=(0, 8))

        # Stream dropdown (Default: 📊 EDC Settlements)
        self.stream_filter = "merchant"
        self.stream_menu = ctk.CTkOptionMenu(
            f_left, values=["📊 EDC Settlements", "🏦 Bank Mutations"],
            width=150, height=28, font=(FONT_FAMILY, 10, "bold"),
            fg_color=WHITE, text_color=TEXT, button_color="#E2E8F0", button_hover_color="#CBD5E1",
            dropdown_font=(FONT_FAMILY, 10, "bold"), command=self._on_stream_changed
        )
        self.stream_menu.pack(side="left", padx=(0, 6))

        # Bank dropdown (Default: ALL Banks)
        self.bank_filter = "ALL"
        self.period_filter = "3d"

        self.bank_menu = ctk.CTkOptionMenu(
            f_left, values=["ALL Banks", "BCA", "MANDIRI", "BRI"],
            width=120, height=28, font=(FONT_FAMILY, 10, "bold"),
            fg_color=WHITE, text_color=TEXT, button_color="#E2E8F0", button_hover_color="#CBD5E1",
            dropdown_font=(FONT_FAMILY, 10, "bold"), command=self._on_bank_changed
        )
        self.bank_menu.pack(side="left", padx=(0, 6))

        # Period dropdown (Default: Last 3 Days, Custom Range max 30 days)
        self.custom_from = ""
        self.custom_to = ""

        self.period_menu = ctk.CTkOptionMenu(
            f_left, values=["Last 3 Days", "Last 7 Days", "Last 14 Days", "Last 30 Days", "Custom Range"],
            width=130, height=28, font=(FONT_FAMILY, 10, "bold"),
            fg_color=WHITE, text_color=TEXT, button_color="#E2E8F0", button_hover_color="#CBD5E1",
            dropdown_font=(FONT_FAMILY, 10, "bold"), command=self._on_period_changed
        )
        self.period_menu.set("Last 3 Days")
        self.period_menu.pack(side="left")

        # Inline Custom Date Range Box (Shown right next to select period when 'Custom Range' is active)
        self.custom_box = ctk.CTkFrame(f_left, fg_color="transparent")

        def _add_one_month(d):
            y = d.year + (d.month // 12)
            m = (d.month % 12) + 1
            max_days = calendar.monthrange(y, m)[1]
            return date(y, m, min(d.day, max_days))

        def _sub_one_month(d):
            y = d.year - (1 if d.month == 1 else 0)
            m = 12 if d.month == 1 else (d.month - 1)
            max_days = calendar.monthrange(y, m)[1]
            return date(y, m, min(d.day, max_days))

        def _get_start_dt():
            s = self.start_entry.get().strip() if hasattr(self, "start_entry") else ""
            if s:
                try:
                    return datetime.strptime(s, "%Y-%m-%d").date()
                except Exception:
                    pass
            return None

        def _get_end_dt():
            s = self.end_entry.get().strip() if hasattr(self, "end_entry") else ""
            if s:
                try:
                    return datetime.strptime(s, "%Y-%m-%d").date()
                except Exception:
                    pass
            return None

        self.start_entry = CTkDateInput(
            self.custom_box, date_pattern="yyyy-mm-dd", placeholder_text="YYYY-MM-DD",
            width=118, height=28, default_date=None,
            max_date=lambda: _get_end_dt(),
            min_date=lambda: _sub_one_month(_get_end_dt()) if _get_end_dt() else None
        )
        self.start_entry.pack(side="left", padx=(0, 2))

        ctk.CTkLabel(self.custom_box, text="→", font=(FONT_FAMILY, 10, "bold"), text_color=MUTED, height=28).pack(side="left", padx=3)

        self.end_entry = CTkDateInput(
            self.custom_box, date_pattern="yyyy-mm-dd", placeholder_text="YYYY-MM-DD",
            width=118, height=28, default_date=None,
            min_date=lambda: _get_start_dt(),
            max_date=lambda: _add_one_month(_get_start_dt()) if _get_start_dt() else None
        )
        self.end_entry.pack(side="left", padx=(2, 4))


        self.apply_btn = ctk.CTkButton(
            self.custom_box, text="Apply", width=55, height=28,
            font=(FONT_FAMILY, 10, "bold"), fg_color=ACCENT, text_color=WHITE, corner_radius=6,
            command=self._on_apply_custom_range
        )
        self.apply_btn.pack(side="left")









        # Right side: Sleek Date Coverage Pill
        f_right = ctk.CTkFrame(top_filter_bar, fg_color="transparent")
        f_right.pack(side="right", padx=(0, 10))

        cov_pill = ctk.CTkFrame(f_right, fg_color=PREVIEW_BG, corner_radius=6, border_color=BORDER, border_width=1)
        cov_pill.pack(side="right")

        self.coverage_lbl = ctk.CTkLabel(
            cov_pill, text="🗓️ Span: —",
            font=(FONT_FAMILY, 9, "bold"), text_color=TEXT, fg_color="transparent"
        )
        self.coverage_lbl.pack(padx=10, pady=3)

        # 1. Cloud KPI Hero Grid (4 Cards aligned at 12 + 4 = 16px)
        kpi_grid = ctk.CTkFrame(self, fg_color="transparent")
        kpi_grid.pack(fill="x", padx=12, pady=(0, 6))
        kpi_grid.rowconfigure(0, weight=1)
        for c in range(4):
            kpi_grid.columnconfigure(c, weight=1, uniform="kpi_col")

        def _make_kpi_card(parent, col, icon, title, is_clickable=False):
            card = ctk.CTkFrame(parent, fg_color=PREVIEW_BG, corner_radius=8, border_color=BORDER_DARK, border_width=1, height=100)
            card.grid(row=0, column=col, padx=4, pady=2, sticky="nsew")
            card.pack_propagate(False)

            top_f = ctk.CTkFrame(card, fg_color="transparent")
            top_f.pack(fill="x", padx=12, pady=(6, 2))

            ico_f = ("Segoe UI Emoji", 13) if IS_WINDOWS else (FONT_FAMILY, 13)
            ico_lbl = ctk.CTkLabel(top_f, text=icon, font=ico_f, text_color=ACCENT, fg_color="transparent")
            ico_lbl.pack(side="left")

            t_lbl = ctk.CTkLabel(top_f, text=title, font=(FONT_FAMILY, 10, "bold"), text_color=MUTED, fg_color="transparent")
            t_lbl.pack(side="left", padx=(6, 0))

            hint_lbl = None
            if is_clickable:
                detail_ico_f = ("Segoe UI Emoji", 14) if IS_WINDOWS else (FONT_FAMILY, 14)
                hint_lbl = ctk.CTkButton(
                    top_f, text="🔍", width=28, height=22, font=detail_ico_f,
                    fg_color="transparent", hover_color=BORDER, text_color=ACCENT, corner_radius=6,
                    command=self._open_settlement_audit_modal
                )
                hint_lbl.pack(side="right")


            val_lbl = ctk.CTkLabel(card, text="—", font=(FONT_FAMILY, 16, "bold"), text_color=TEXT, fg_color="transparent", anchor="w")
            val_lbl.pack(fill="x", padx=12, pady=(1, 2))

            sub_lbl = ctk.CTkLabel(card, text="—", font=(FONT_FAMILY, 10, "bold"), text_color=MUTED, fg_color="transparent", anchor="w")
            sub_lbl.pack(fill="x", padx=12, pady=(1, 4))

            if is_clickable:
                clickable_widgets = [card, top_f, ico_lbl, t_lbl, val_lbl, sub_lbl]
                if hint_lbl: clickable_widgets.append(hint_lbl)
                for w in clickable_widgets:
                    try:
                        w.configure(cursor="hand2")
                    except Exception:
                        pass
                    w.bind("<Button-1>", lambda e: self._open_settlement_audit_modal())

            return val_lbl, sub_lbl

        self.kpi_gross_val, self.kpi_gross_sub = _make_kpi_card(kpi_grid, 0, "💰", "SETTLED GROSS REVENUE")
        self.kpi_txns_val,  self.kpi_txns_sub  = _make_kpi_card(kpi_grid, 1, "🧾", "TOTAL TRANSACTIONS", is_clickable=True)
        self.kpi_pace_val,  self.kpi_pace_sub  = _make_kpi_card(kpi_grid, 2, "📈", "DAILY REVENUE PACE")
        self.kpi_sync_val,  self.kpi_sync_sub  = _make_kpi_card(kpi_grid, 3, "👥", "CLOUD SYNC STATUS")


        # 2. Middle Section: 3 Executive Insight Distribution Panels (aligned at 12 + 4 = 16px)
        mid_grid = ctk.CTkFrame(self, fg_color="transparent")
        mid_grid.pack(fill="x", padx=12, pady=(0, 6))
        mid_grid.columnconfigure(0, weight=1, uniform="mid_col")
        mid_grid.columnconfigure(1, weight=1, uniform="mid_col")
        mid_grid.columnconfigure(2, weight=1, uniform="mid_col")

        def _make_insight_panel(parent, col, title, icon):
            panel = ctk.CTkFrame(parent, fg_color=PANEL, corner_radius=8, border_color=BORDER_DARK, border_width=1, height=170)
            panel.grid(row=0, column=col, padx=4, pady=2, sticky="nsew")
            panel.pack_propagate(False)

            hdr = ctk.CTkFrame(panel, fg_color=PREVIEW_BG, corner_radius=6, height=28)
            hdr.pack(fill="x", padx=8, pady=(6, 4))
            hdr.pack_propagate(False)

            ico_f = ("Segoe UI Emoji", 11) if IS_WINDOWS else (FONT_FAMILY, 11)
            ctk.CTkLabel(hdr, text=icon, font=ico_f, text_color=ACCENT, fg_color="transparent").pack(side="left", padx=(8, 4))
            ctk.CTkLabel(hdr, text=title, font=(FONT_FAMILY, 10, "bold"), text_color=TEXT, fg_color="transparent").pack(side="left", padx=(2, 6))

            content_box = ctk.CTkFrame(panel, fg_color="transparent")
            content_box.pack(fill="both", expand=True, padx=8, pady=(0, 6))
            return content_box

        self.panel_stores = _make_insight_panel(mid_grid, 0, "OUTLET & STORE SHARE", "🏪")
        self.panel_banks  = _make_insight_panel(mid_grid, 1, "BANK SETTLEMENT SPLIT", "🏦")
        self.panel_cards  = _make_insight_panel(mid_grid, 2, "TRADING PERIOD MIX", "📅")







        # 3. Bottom Section: Responsive Canvas Line Chart (Aligned with 16px Grid Edge)
        daily_card = ctk.CTkFrame(self, fg_color=PANEL, corner_radius=8, border_color=BORDER_DARK, border_width=1)
        daily_card.pack(fill="both", expand=True, padx=16, pady=(0, 6))


        daily_hdr = ctk.CTkFrame(daily_card, fg_color="transparent")
        daily_hdr.pack(fill="x", padx=14, pady=(6, 2))

        ctk.CTkLabel(
            daily_hdr, text="📈 Revenue Velocity & Daily Performance Trend",
            font=(FONT_FAMILY, 11, "bold"), text_color=TEXT, fg_color="transparent"
        ).pack(side="left")

        self.daily_metrics_box = ctk.CTkFrame(daily_hdr, fg_color="transparent")
        self.daily_metrics_box.pack(side="right")

        self.timeline_pace_lbl = ctk.CTkLabel(self.daily_metrics_box, text="", font=(FONT_FAMILY, 10, "bold"), text_color=MUTED, fg_color="transparent")
        self.timeline_pace_lbl.pack(side="right")

        chart_frame = ctk.CTkFrame(daily_card, fg_color=PREVIEW_BG, corner_radius=6)
        chart_frame.pack(fill="both", expand=True, padx=8, pady=(0, 6))

        self.chart_canvas = tk.Canvas(chart_frame, bg=PREVIEW_BG, highlightthickness=0)
        self.chart_canvas.pack(fill="both", expand=True, padx=4, pady=4)

        self._cached_daily_stats = []
        self._cached_granularity = "daily"
        self._chart_points = []
        self._hover_idx = None

        self.chart_canvas.bind("<Configure>", self._on_canvas_resize)
        self.chart_canvas.bind("<Motion>", self._on_canvas_motion)
        self.chart_canvas.bind("<Leave>", self._on_canvas_leave)

    def _on_stream_changed(self, choice: str):
        self.stream_filter = "mutation" if "Mutation" in choice else "merchant"
        self.update_summary()

    def _on_bank_changed(self, choice: str):
        self.bank_filter = "ALL" if choice == "ALL Banks" else choice
        self.update_summary()

    def _on_period_changed(self, choice: str):
        if choice == "Custom Range":
            self.custom_box.pack(side="left", padx=(6, 0))
            return

        self.custom_box.pack_forget()
        self.custom_from = ""
        self.custom_to = ""

        mapping = {
            "Last 3 Days": "3d",
            "Last 7 Days": "7d",
            "Last 14 Days": "14d",
            "Last 30 Days": "30d",
        }
        self.period_filter = mapping.get(choice, "3d")
        self.update_summary()

    def _on_apply_custom_range(self):
        f_str = self.start_entry.get().strip()
        t_str = self.end_entry.get().strip()
        if not f_str or not t_str:
            messagebox.showwarning("Missing Date", "Please enter both Start Date and End Date.", parent=self.winfo_toplevel())
            return
        try:
            f_dt = datetime.strptime(f_str, "%Y-%m-%d")
            t_dt = datetime.strptime(t_str, "%Y-%m-%d")
        except Exception:
            messagebox.showwarning("Invalid Date Format", "Please enter dates in YYYY-MM-DD format (e.g. 2026-07-06).", parent=self.winfo_toplevel())
            return

        if f_dt > t_dt:
            messagebox.showwarning("Invalid Date Range", "Start Date cannot be after End Date.", parent=self.winfo_toplevel())
            return

        # Check if t_dt exceeds 1 calendar month from f_dt
        y = f_dt.year + (f_dt.month // 12)
        m = (f_dt.month % 12) + 1
        max_d = min(f_dt.day, calendar.monthrange(y, m)[1])
        max_allowed = date(y, m, max_d)
        if t_dt.date() > max_allowed:
            messagebox.showwarning(
                "Maximum 1 Month Limit",
                f"Maximum allowed date range is 1 month.\nYour selected range ends on {t_str}, but the maximum allowed end date for {f_str} is {max_allowed.strftime('%Y-%m-%d')}.\n\nPlease select a range within 1 month.",
                parent=self.winfo_toplevel()
            )
            return

        self.custom_from = f_dt.strftime("%Y-%m-%d")
        self.custom_to = t_dt.strftime("%Y-%m-%d")
        self.period_filter = "custom"
        self.update_summary()




    def update_summary(self):
        """Fetch executive financial analytics in background."""
        def _bg():
            try:
                import cloud_sync
                stats = cloud_sync.fetch_cloud_analytics(
                    bank=self.bank_filter,
                    period=self.period_filter,
                    custom_from=self.custom_from,
                    custom_to=self.custom_to,
                    data_type=self.stream_filter
                )
                self.after(0, lambda: self._apply_analytics(stats))
            except Exception as e:
                print(f"[CloudDashboard] Analytics update error: {e}")

        threading.Thread(target=_bg, daemon=True).start()


    def query_transactions(self):
        """Alias for refreshing analytics."""
        self.update_summary()

    def _open_settlement_audit_modal(self):
        stats = getattr(self, "_last_stats", None)
        if not stats:
            return

        top = ctk.CTkToplevel(self)
        top.title("Bank Settlement & Missing Dates Audit")
        top.geometry("880x700")
        top.minsize(780, 560)
        top.configure(fg_color=PANEL)
        top.transient(self.winfo_toplevel())

        # Top Header Banner
        hdr_frame = ctk.CTkFrame(top, fg_color=WHITE, corner_radius=8, border_color=BORDER_DARK, border_width=1)
        hdr_frame.pack(fill="x", padx=16, pady=(16, 8))

        h_left = ctk.CTkFrame(hdr_frame, fg_color="transparent")
        h_left.pack(side="left", padx=14, pady=12)

        stream_lbl = "Bank Mutations" if getattr(self, "stream_filter", "merchant") == "mutation" else "EDC Settlements"
        ctk.CTkLabel(
            h_left, text=f"🧾 {stream_lbl} & Missing Dates Audit",
            font=(FONT_FAMILY, 15, "bold"), text_color=TEXT
        ).pack(anchor="w")

        bank_name = self.bank_filter
        span_str = stats.get("date_span", "—")
        scope_str = f"Stream: {stream_lbl} • Bank: {bank_name} • Period: {span_str}"
        ctk.CTkLabel(
            h_left, text=scope_str,
            font=(FONT_FAMILY, 11, "bold"), text_color=MUTED
        ).pack(anchor="w", pady=(3, 0))

        # KPI Summary Chips on the Right
        h_right = ctk.CTkFrame(hdr_frame, fg_color="transparent")
        h_right.pack(side="right", padx=14, pady=12)

        active_cnt = stats.get("active_days_count", 0)
        missing_cnt = stats.get("missing_dates_count", 0)
        expected_cnt = stats.get("expected_days_count", active_cnt + missing_cnt)

        chip1 = ctk.CTkFrame(h_right, fg_color=PREVIEW_BG, corner_radius=6, border_color=BORDER, border_width=1)
        chip1.pack(side="left", padx=4)
        ctk.CTkLabel(chip1, text=f"✅ {active_cnt}/{expected_cnt} Days Settled", font=(FONT_FAMILY, 11, "bold"), text_color=SUCCESS).pack(padx=10, pady=5)

        if missing_cnt > 0:
            chip2 = ctk.CTkFrame(h_right, fg_color="#FEF2F2", corner_radius=6, border_color="#FECACA", border_width=1)
            chip2.pack(side="left", padx=4)
            ctk.CTkLabel(chip2, text=f"⚠️ {missing_cnt} Days Missing", font=(FONT_FAMILY, 11, "bold"), text_color="#EF4444").pack(padx=10, pady=5)

        # Bank Breakdown Cards (BCA, MANDIRI, BRI)
        bank_grid = ctk.CTkFrame(top, fg_color="transparent")
        bank_grid.pack(fill="x", padx=12, pady=(0, 6))

        banks_data = stats.get("by_bank", {})
        all_banks = ["BCA", "MANDIRI", "BRI"] if self.bank_filter == "ALL" else [self.bank_filter]
        bank_colors = {"BCA": "#0066AE", "MANDIRI": "#D97706", "BRI": "#004B87"}

        daily_list = stats.get("daily_stats", [])
        daily_map = {d["date"]: d for d in daily_list}
        missing_list = stats.get("missing_dates", [])

        # Gather calendar days
        all_dates = sorted(list(set(list(daily_map.keys()) + missing_list)))

        for b_name in all_banks:
            b_col = bank_colors.get(b_name.upper(), ACCENT)
            b_active_dates = [d for d in all_dates if d in daily_map and b_name in daily_map[d].get("banks", {})]
            b_active_cnt = len(b_active_dates)
            b_miss_cnt = len(all_dates) - b_active_cnt
            b_gross = banks_data.get(b_name, {}).get("gross", 0.0)
            b_txns = banks_data.get(b_name, {}).get("count", 0)

            b_card = ctk.CTkFrame(bank_grid, fg_color=WHITE, corner_radius=6, border_color=BORDER_DARK, border_width=1)
            b_card.pack(side="left", fill="both", expand=True, padx=4)

            b_head = ctk.CTkFrame(b_card, fg_color="transparent")
            b_head.pack(fill="x", padx=10, pady=(8, 2))
            ctk.CTkLabel(b_head, text=b_name, font=(FONT_FAMILY, 13, "bold"), text_color=b_col).pack(side="left")
            ctk.CTkLabel(b_head, text=f"Rp {b_gross/1_000_000:,.1f}M ({b_txns:,} txns)", font=(FONT_FAMILY, 10, "bold"), text_color=MUTED).pack(side="right")

            stat_txt = f"{b_active_cnt} Settled" + (f" • ⚠️ {b_miss_cnt} Missing" if b_miss_cnt > 0 else " • 100% Complete")
            stat_col = "#D97706" if b_miss_cnt > 0 else SUCCESS
            ctk.CTkLabel(b_card, text=stat_txt, font=(FONT_FAMILY, 11, "bold"), text_color=stat_col).pack(anchor="w", padx=10, pady=(0, 8))

        # Filter Tabs Bar
        filter_bar = ctk.CTkFrame(top, fg_color="transparent")
        filter_bar.pack(fill="x", padx=16, pady=(4, 6))

        tab_var = tk.StringVar(value="all")

        # Frozen Column Header Container (Fixed at top, outside scroll area)
        th_container = ctk.CTkFrame(top, fg_color=PREVIEW_BG, corner_radius=6, border_color=BORDER_DARK, border_width=1, height=36)
        th_container.pack(fill="x", padx=16, pady=(0, 2))

        # Table Rows Container (Scrollable rows only)
        table_frame = ctk.CTkScrollableFrame(top, fg_color=WHITE, corner_radius=6, border_color=BORDER_DARK, border_width=1)
        table_frame.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        def _render_matrix():
            for w in th_container.winfo_children():
                w.destroy()
            for w in table_frame.winfo_children():
                w.destroy()

            # Frozen Matrix Header
            ctk.CTkLabel(th_container, text="Bank Account", width=180, font=(FONT_FAMILY, 11, "bold"), text_color=TEXT, anchor="w").pack(side="left", padx=12, pady=6)
            ctk.CTkLabel(th_container, text="Merchant Settlements Date Range", width=280, font=(FONT_FAMILY, 11, "bold"), text_color=TEXT, anchor="w").pack(side="left", padx=8, pady=6)
            ctk.CTkLabel(th_container, text="Bank Mutations Date Range", width=280, font=(FONT_FAMILY, 11, "bold"), text_color=TEXT, anchor="w").pack(side="left", padx=8, pady=6)

            try:
                import cloud_sync
                matrix_data = cloud_sync.fetch_cloud_coverage_matrix()
            except Exception:
                matrix_data = []

            if not matrix_data:
                ctk.CTkLabel(table_frame, text="No coverage data found", font=(FONT_FAMILY, 11, "bold"), text_color=MUTED).pack(pady=40)
                return

            for row in matrix_data:
                b_name = row.get("bank", "")
                m_dates = row.get("merchant_dates", "—")
                mut_dates = row.get("mutation_dates", "—")
                m_col = bank_colors.get(b_name.upper(), TEXT)

                tr = ctk.CTkFrame(table_frame, fg_color="transparent", corner_radius=4)
                tr.pack(fill="x", pady=4)

                ctk.CTkLabel(tr, text=b_name, width=180, font=(FONT_FAMILY, 11, "bold"), text_color=m_col, anchor="w").pack(side="left", padx=12, pady=6)
                ctk.CTkLabel(tr, text=m_dates, width=280, font=(FONT_FAMILY, 11, "bold" if m_dates != "—" else "normal"), text_color=SUCCESS if m_dates != "—" else MUTED, anchor="w").pack(side="left", padx=8)
                ctk.CTkLabel(tr, text=mut_dates, width=280, font=(FONT_FAMILY, 11, "bold" if mut_dates != "—" else "normal"), text_color=SUCCESS if mut_dates != "—" else MUTED, anchor="w").pack(side="left", padx=8)

        def _render_table(filter_mode):
            if filter_mode == "matrix":
                _render_matrix()
                return

            for w in th_container.winfo_children():
                w.destroy()
            for w in table_frame.winfo_children():
                w.destroy()

            # Frozen Table Header
            ctk.CTkLabel(th_container, text="Settlement Date", width=160, font=(FONT_FAMILY, 11, "bold"), text_color=TEXT, anchor="w").pack(side="left", padx=8, pady=6)
            ctk.CTkLabel(th_container, text="Status", width=120, font=(FONT_FAMILY, 11, "bold"), text_color=TEXT, anchor="w").pack(side="left", padx=4, pady=6)
            for b_name in all_banks:
                ctk.CTkLabel(th_container, text=b_name, width=120, font=(FONT_FAMILY, 11, "bold"), text_color=bank_colors.get(b_name, TEXT), anchor="e").pack(side="left", padx=4, pady=6)
            ctk.CTkLabel(th_container, text="Total Settled", width=170, font=(FONT_FAMILY, 11, "bold"), text_color=TEXT, anchor="e").pack(side="left", padx=8, pady=6)

            filtered_dates = []
            for d in all_dates:
                is_missing = (d in missing_list) or (d not in daily_map)
                if filter_mode == "missing" and not is_missing:
                    continue
                if filter_mode == "settled" and is_missing:
                    continue
                filtered_dates.append((d, is_missing))

            if not filtered_dates:
                ctk.CTkLabel(table_frame, text="No dates matching filter", font=(FONT_FAMILY, 11, "bold"), text_color=MUTED).pack(pady=40)
                return

            for d_str, is_missing in filtered_dates:
                info = daily_map.get(d_str)
                row_bg = "#FEF2F2" if is_missing else "transparent"
                row_border = "#FECACA" if is_missing else BORDER

                tr = ctk.CTkFrame(table_frame, fg_color=row_bg, corner_radius=4, border_color=row_border, border_width=1 if is_missing else 0)
                tr.pack(fill="x", pady=2)

                try:
                    dt_obj = datetime.strptime(d_str, "%Y-%m-%d")
                    d_label = dt_obj.strftime("%d %b %Y (%a)")
                except Exception:
                    d_label = d_str

                ctk.CTkLabel(tr, text=d_label, width=160, font=(FONT_FAMILY, 11, "bold"), text_color=TEXT if not is_missing else "#991B1B", anchor="w").pack(side="left", padx=8, pady=6)

                status_txt = "✅ Settled" if not is_missing else "⚠️ Missing"
                status_col = SUCCESS if not is_missing else "#EF4444"
                ctk.CTkLabel(tr, text=status_txt, width=120, font=(FONT_FAMILY, 11, "bold"), text_color=status_col, anchor="w").pack(side="left", padx=4)

                for b_name in all_banks:
                    b_vol = info.get("banks", {}).get(b_name, 0.0) if info else 0.0
                    b_txt = f"Rp {b_vol/1_000_000:,.1f}M" if b_vol > 0 else ("—" if not is_missing else "Missing")
                    b_col = TEXT if b_vol > 0 else ("#94A3B8" if not is_missing else "#EF4444")
                    ctk.CTkLabel(tr, text=b_txt, width=120, font=(FONT_FAMILY, 11, "bold" if b_vol > 0 else "normal"), text_color=b_col, anchor="e").pack(side="left", padx=4)

                g_val = info.get("gross", 0.0) if info else 0.0
                c_val = info.get("count", 0) if info else 0
                tot_txt = f"Rp {g_val/1_000_000:,.1f}M ({c_val} txns)" if g_val > 0 else "Rp 0 (0 txns)"
                tot_col = TEXT if g_val > 0 else "#94A3B8"
                ctk.CTkLabel(tr, text=tot_txt, width=170, font=(FONT_FAMILY, 11, "bold" if g_val > 0 else "normal"), text_color=tot_col, anchor="e").pack(side="left", padx=8)

        # Tab Buttons
        def _set_tab(mode):
            tab_var.set(mode)
            btn_all.configure(fg_color=ACCENT if mode == "all" else "transparent", text_color=WHITE if mode == "all" else TEXT)
            btn_miss.configure(fg_color="#EF4444" if mode == "missing" else "transparent", text_color=WHITE if mode == "missing" else TEXT)
            btn_sett.configure(fg_color=SUCCESS if mode == "settled" else "transparent", text_color=WHITE if mode == "settled" else TEXT)
            btn_matrix.configure(fg_color=ACCENT if mode == "matrix" else "transparent", text_color=WHITE if mode == "matrix" else TEXT)
            _render_table(mode)

        btn_all = ctk.CTkButton(filter_bar, text=f"All Dates ({len(all_dates)})", width=110, height=32, font=(FONT_FAMILY, 11, "bold"), fg_color=ACCENT, text_color=WHITE, corner_radius=6, command=lambda: _set_tab("all"))
        btn_all.pack(side="left", padx=(0, 4))

        btn_miss = ctk.CTkButton(filter_bar, text=f"⚠️ Missing Dates ({missing_cnt})", width=140, height=32, font=(FONT_FAMILY, 11, "bold"), fg_color="transparent", text_color=TEXT, border_color=BORDER_DARK, border_width=1, corner_radius=6, command=lambda: _set_tab("missing"))
        btn_miss.pack(side="left", padx=4)

        btn_sett = ctk.CTkButton(filter_bar, text=f"✅ Settled Dates ({active_cnt})", width=135, height=32, font=(FONT_FAMILY, 11, "bold"), fg_color="transparent", text_color=TEXT, border_color=BORDER_DARK, border_width=1, corner_radius=6, command=lambda: _set_tab("settled"))
        btn_sett.pack(side="left", padx=4)

        btn_matrix = ctk.CTkButton(filter_bar, text="📅 Date Coverage Matrix", width=160, height=32, font=(FONT_FAMILY, 11, "bold"), fg_color="transparent", text_color=TEXT, border_color=BORDER_DARK, border_width=1, corner_radius=6, command=lambda: _set_tab("matrix"))
        btn_matrix.pack(side="left", padx=4)

        # Initial render
        _render_table("all")

        # Bottom Close Button
        btn_close = ctk.CTkButton(
            top, text="Close Audit Window", font=(FONT_FAMILY, 12, "bold"),
            fg_color=ACCENT, text_color=WHITE, height=38, width=170, corner_radius=6,
            command=top.destroy
        )
        btn_close.pack(pady=(0, 14))

        top.grab_set()



    def _apply_analytics(self, stats: dict):
        self._last_stats = stats
        # Always preserve latest Cloud Sync Status regardless of filter
        if stats.get("last_uploader") and stats["last_uploader"] != "—":
            self.kpi_sync_val.configure(text=stats["last_uploader"], text_color=SUCCESS)
            last_up = stats.get("last_updated", "—")
            self.kpi_sync_sub.configure(text=f"Last Synced: {last_up}", text_color=MUTED)

        missing_cnt = stats.get("missing_dates_count", 0)
        missing_list = stats.get("missing_dates", [])
        date_span = stats.get("date_span", "—")


        if not stats.get("configured") or stats.get("total_txns", 0) == 0:
            self.kpi_gross_val.configure(text="Rp 0", text_color=MUTED)
            self.kpi_gross_sub.configure(text="No Settled Sales", text_color=MUTED)
            self.kpi_txns_val.configure(text="0 Transactions", text_color=MUTED)

            if missing_cnt > 0:
                try:
                    m_sample = [datetime.strptime(d, "%Y-%m-%d").strftime("%d %b") for d in missing_list[:3]]
                    m_str = ", ".join(m_sample) + ("..." if len(missing_list) > 3 else "")
                except Exception:
                    m_str = f"{missing_cnt} dates"
                self.kpi_txns_sub.configure(text=f"⚠️ {missing_cnt} Missing Settlement Days", text_color="#D97706")
                self.coverage_lbl.configure(text=f"Span: {date_span} (⚠️ {missing_cnt} Missing Days: {m_str})", text_color="#D97706")
            else:
                self.kpi_txns_sub.configure(text="No data for filter", text_color=MUTED)
                self.coverage_lbl.configure(text=f"Span: {date_span}" if date_span != "—" else "Span: No Records", text_color=MUTED)

            self.kpi_pace_val.configure(text="Rp 0.0M", text_color=MUTED)
            self.kpi_pace_sub.configure(text="0 active dates", text_color=MUTED)
            self.timeline_pace_lbl.configure(text="⚡ Pace: Rp 0.0M / day")

            for p in (self.panel_stores, self.panel_banks, self.panel_cards):
                for w in p.winfo_children(): w.destroy()
                ctk.CTkLabel(p, text="No Data for Period", font=(FONT_FAMILY, 9), text_color=MUTED).pack(pady=20)

            self.chart_canvas.delete("all")
            cw = max(200, self.chart_canvas.winfo_width())
            ch = max(100, self.chart_canvas.winfo_height())

            if missing_cnt > 0:
                missing_full_str = ", ".join(missing_list[:6]) + (f" (+{len(missing_list)-6} more)" if len(missing_list) > 6 else "")
                self.chart_canvas.create_text(
                    cw / 2, ch / 2 - 14,
                    text=f"⚠️ Missing Settlement Data for {date_span}",
                    fill="#D97706", font=(FONT_FAMILY, 11, "bold")
                )
                self.chart_canvas.create_text(
                    cw / 2, ch / 2 + 12,
                    text=f"No settlement reports uploaded for {missing_cnt} expected calendar days ({missing_full_str})",
                    fill="#64748B", font=(FONT_FAMILY, 9)
                )
            else:
                self.chart_canvas.create_text(cw / 2, ch / 2, text="No transactions found for the selected filter.", fill="#94A3B8", font=(FONT_FAMILY, 10, "bold"))
            return



        # 1. KPI Top Cards
        gross = stats["total_gross"]
        txns = stats["total_txns"]

        pace = stats["daily_run_rate"]
        granularity = stats.get("granularity", "daily")
        unit_str = "mo" if granularity == "monthly" else "day"
        days_cnt = stats.get("active_days_count", 0)
        missing_cnt = stats.get("missing_dates_count", 0)
        missing_list = stats.get("missing_dates", [])
        is_mutation = stats.get("data_type") == "mutation"
        self.kpi_gross_val.configure(text=f"Rp {gross:,.0f}".replace(",", "."), text_color=TEXT)
        self.kpi_gross_sub.configure(text="Bank Mutation Volume" if is_mutation else "Settled Sales Revenue", text_color=MUTED)

        self.kpi_txns_val.configure(text=f"{txns:,} Transactions", text_color=TEXT)
        stream_name = "Mutation" if is_mutation else "Settlement"
        if missing_cnt > 0:
            self.kpi_txns_sub.configure(text=f"Across {days_cnt} Days (⚠️ {missing_cnt} Missing Days)", text_color="#D97706")
        else:
            self.kpi_txns_sub.configure(text=f"Across {days_cnt} {stream_name} Days", text_color=MUTED)

        self.kpi_pace_val.configure(text=f"Rp {pace/1_000_000:,.1f}M", text_color=TEXT)
        self.kpi_pace_sub.configure(text=f"Velocity per active {unit_str}", text_color=MUTED)

        self.kpi_sync_val.configure(text=stats.get("last_uploader", "—"), text_color=SUCCESS)
        last_up = stats.get("last_updated", "—")
        self.kpi_sync_sub.configure(text=f"Last Synced: {last_up}", text_color=MUTED)

        # 2. Outlet & Store Share
        for w in self.panel_stores.winfo_children(): w.destroy()
        stores_data = stats.get("by_store", {})
        if stores_data:
            s_sorted = sorted(stores_data.items(), key=lambda x: x[1]["gross"], reverse=True)
            for s_name, s_info in s_sorted[:4]:
                s_pct = (s_info["gross"] / gross * 100) if gross > 0 else 0
                s_row = ctk.CTkFrame(self.panel_stores, fg_color="transparent")
                s_row.pack(fill="x", pady=2)

                s_head = ctk.CTkFrame(s_row, fg_color="transparent")
                s_head.pack(fill="x")
                ctk.CTkLabel(s_head, text=s_name, font=(FONT_FAMILY, 9, "bold"), text_color=TEXT).pack(side="left")
                ctk.CTkLabel(s_head, text=f"{s_pct:.1f}%  (Rp {s_info['gross']/1_000_000:,.1f}M)", font=(FONT_FAMILY, 8), text_color=MUTED).pack(side="right")

                pb = ctk.CTkProgressBar(s_row, height=5, corner_radius=3, progress_color="#6366F1", fg_color=BORDER)
                pb.pack(fill="x", pady=(2, 0))
                pb.set(max(0.02, s_pct / 100.0))
        else:
            ctk.CTkLabel(self.panel_stores, text="No Store Data", font=(FONT_FAMILY, 9), text_color=MUTED).pack(pady=20)

        # 3. Bank Settlement Split
        for w in self.panel_banks.winfo_children(): w.destroy()
        banks_data = stats.get("by_bank", {})
        if banks_data:
            b_sorted = sorted(banks_data.items(), key=lambda x: x[1]["gross"], reverse=True)
            bank_colors = {"BCA": "#0066AE", "MANDIRI": "#D97706", "BRI": "#004B87"}
            for b_name, b_info in b_sorted[:4]:
                b_pct = (b_info["gross"] / gross * 100) if gross > 0 else 0
                b_col = bank_colors.get(b_name.upper(), ACCENT)

                b_row = ctk.CTkFrame(self.panel_banks, fg_color="transparent")
                b_row.pack(fill="x", pady=2)

                b_head = ctk.CTkFrame(b_row, fg_color="transparent")
                b_head.pack(fill="x")
                ctk.CTkLabel(b_head, text=b_name, font=(FONT_FAMILY, 9, "bold"), text_color=b_col).pack(side="left")
                ctk.CTkLabel(b_head, text=f"{b_pct:.1f}%  (Rp {b_info['gross']/1_000_000:,.1f}M)", font=(FONT_FAMILY, 8), text_color=MUTED).pack(side="right")

                pb = ctk.CTkProgressBar(b_row, height=5, corner_radius=3, progress_color=b_col, fg_color=BORDER)
                pb.pack(fill="x", pady=(2, 0))
                pb.set(max(0.02, b_pct / 100.0))
        else:
            ctk.CTkLabel(self.panel_banks, text="No Bank Data", font=(FONT_FAMILY, 9), text_color=MUTED).pack(pady=20)

        # 4. Weekday, Weekend & Long Weekend Trading Performance
        for w in self.panel_cards.winfo_children(): w.destroy()

        daily_list = stats.get("daily_stats", [])

        # Load official Indonesian holiday calendar
        try:
            import holidays
            id_holidays = holidays.Indonesia(years=list(range(2023, 2030)))
        except Exception:
            id_holidays = {}

        def _classify_date(dt_obj):
            d_str = dt_obj.strftime("%Y-%m-%d")
            d_date = dt_obj.date()
            # 1. Long Weekend = Official Indonesian Holiday from holidays library
            if (d_str in id_holidays) or (d_date in id_holidays):
                return "Long Weekend"
            # 2. Weekend = Saturday & Sunday
            if dt_obj.weekday() in (5, 6):
                return "Weekend"
            # 3. Weekday = Monday through Friday
            return "Weekday"

        period_data = {
            "Weekday": {"gross": 0.0, "count": 0, "color": "#6366F1", "icon": "💼"},
            "Weekend": {"gross": 0.0, "count": 0, "color": "#10B981", "icon": "🏖️"},
            "Long Weekend": {"gross": 0.0, "count": 0, "color": "#F59E0B", "icon": "🎉"},
        }


        for d in daily_list:
            d_str = d.get("date", "")
            g_val = float(d.get("gross", 0.0))
            c_val = int(d.get("count", 0))
            try:
                dt = datetime.strptime(d_str, "%Y-%m-%d")
                cat = _classify_date(dt)
            except Exception:
                cat = "Weekday"
            period_data[cat]["gross"] += g_val
            period_data[cat]["count"] += c_val

        for cat_name, info in period_data.items():
            cat_g = info["gross"]
            cat_pct = (cat_g / gross * 100) if gross > 0 else 0.0
            cat_col = info["color"]
            cat_ico = info["icon"]

            m_row = ctk.CTkFrame(self.panel_cards, fg_color="transparent")
            m_row.pack(fill="x", pady=3)

            m_head = ctk.CTkFrame(m_row, fg_color="transparent")
            m_head.pack(fill="x")
            ctk.CTkLabel(m_head, text=f"{cat_ico} {cat_name}", font=(FONT_FAMILY, 9, "bold"), text_color=TEXT).pack(side="left")
            ctk.CTkLabel(m_head, text=f"{cat_pct:.1f}%  (Rp {cat_g/1_000_000:,.1f}M)", font=(FONT_FAMILY, 8), text_color=MUTED).pack(side="right")

            pb = ctk.CTkProgressBar(m_row, height=5, corner_radius=3, progress_color=cat_col, fg_color=BORDER)
            pb.pack(fill="x", pady=(2, 0))
            pb.set(max(0.01 if cat_pct > 0 else 0.0, cat_pct / 100.0))



        # 5. Cache stats and trigger Responsive Canvas Line Chart with Missing Dates Info
        if missing_cnt > 0:
            try:
                m_sample = [datetime.strptime(d, "%Y-%m-%d").strftime("%d %b") for d in missing_list[:3]]
                m_str = ", ".join(m_sample) + ("..." if len(missing_list) > 3 else "")
            except Exception:
                m_str = f"{missing_cnt} dates"
            self.coverage_lbl.configure(
                text=f"Span: {stats['date_span']} ({days_cnt} Days • ⚠️ {missing_cnt} Missing Days: {m_str})",
                text_color="#D97706"
            )
        else:
            self.coverage_lbl.configure(
                text=f"Span: {stats['date_span']} ({days_cnt} {'Months' if granularity == 'monthly' else 'Days'} • 100% Full)",
                text_color=TEXT
            )

        self.timeline_pace_lbl.configure(text=f"⚡ Pace: Rp {pace/1_000_000:,.1f}M / {unit_str}")

        self._cached_daily_stats = stats.get("daily_stats", [])
        self._cached_granularity = granularity
        self._render_line_chart()

    def _on_canvas_resize(self, event):
        self._render_line_chart()

    def _render_line_chart(self):
        self.chart_canvas.delete("all")
        if not self._cached_daily_stats:
            return

        w = self.chart_canvas.winfo_width()
        h = self.chart_canvas.winfo_height()
        if w < 100 or h < 80:
            return

        daily_list = sorted(self._cached_daily_stats, key=lambda x: x["date"])
        n = len(daily_list)
        if n == 0:
            return

        max_daily_g = max((d["gross"] for d in daily_list), default=1.0) * 1.18
        if max_daily_g <= 0: max_daily_g = 1.0

        pad_l, pad_r, pad_t, pad_b = 65, 35, 35, 30
        plot_w = max(40, w - pad_l - pad_r)
        plot_h = max(40, h - pad_t - pad_b)

        # 1. Y-Axis Gridlines & Left-Side Revenue Ticks
        for i in range(5):
            val = max_daily_g * (i / 4.0)
            y = pad_t + plot_h - (i / 4.0) * plot_h
            self.chart_canvas.create_line(pad_l, y, w - pad_r, y, fill="#E2E8F0", dash=(2, 3), tags="grid")
            vol_s = f"{val/1_000_000:,.0f}M" if val >= 1_000_000 else f"{val:,.0f}"
            self.chart_canvas.create_text(pad_l - 8, y, text=vol_s, anchor="e", fill="#94A3B8", font=(FONT_FAMILY, 8, "bold"), tags="grid")

        # 2. Compute Data Coordinates (1x canvas coordinate scale)
        points = []
        for idx, d_info in enumerate(daily_list):
            d_str = d_info["date"]
            d_g = d_info["gross"]
            d_cnt = d_info["count"]
            banks_dict = d_info.get("banks", {})

            x = pad_l + (idx / (n - 1)) * plot_w if n > 1 else pad_l + plot_w / 2
            y = pad_t + plot_h - (d_g / max_daily_g) * plot_h

            display_date = d_str
            try:
                if self._cached_granularity == "monthly":
                    dt_obj = datetime.strptime(d_str, "%Y-%m")
                    display_date = dt_obj.strftime("%b %Y")
                else:
                    dt_obj = datetime.strptime(d_str, "%Y-%m-%d")
                    display_date = dt_obj.strftime("%d %b")
            except Exception:
                pass

            points.append((x, y, display_date, d_g, d_cnt, banks_dict, d_str))
        self._chart_points = points

        # 3. Compute High-Density Smooth Bezier Curve (Tension 0.28, 32 steps per interval)
        tension = 0.28
        num_steps = 32
        curve_coords = []

        if len(points) < 2:
            curve_coords = [c for p in points for c in (p[0], p[1])]
        elif len(points) == 2:
            curve_coords = [points[0][0], points[0][1], points[1][0], points[1][1]]
        else:
            n_pts = len(points)
            pts = [(float(p[0]), float(p[1])) for p in points]

            for i in range(n_pts - 1):
                p0 = pts[max(0, i - 1)]
                p1 = pts[i]
                p2 = pts[i + 1]
                p3 = pts[min(n_pts - 1, i + 2)]

                cp1x = p1[0] + (p2[0] - p0[0]) * (tension / 2.0)
                cp1y = p1[1] + (p2[1] - p0[1]) * (tension / 2.0)
                cp2x = p2[0] - (p3[0] - p1[0]) * (tension / 2.0)
                cp2y = p2[1] - (p3[1] - p1[1]) * (tension / 2.0)

                for step in range(num_steps):
                    t = step / float(num_steps)
                    mt = 1.0 - t
                    bx = (mt**3) * p1[0] + 3.0 * (mt**2) * t * cp1x + 3.0 * mt * (t**2) * cp2x + (t**3) * p2[0]
                    by = (mt**3) * p1[1] + 3.0 * (mt**2) * t * cp1y + 3.0 * mt * (t**2) * cp2y + (t**3) * p2[1]
                    curve_coords.extend([bx, by])

            curve_coords.extend([pts[-1][0], pts[-1][1]])

        # 4. Soft Gradient Area Fill & Smooth Line Stroke
        if len(points) >= 2 and len(curve_coords) >= 4:
            poly = [pad_l, pad_t + plot_h] + curve_coords + [pad_l + plot_w, pad_t + plot_h]
            self.chart_canvas.create_polygon(poly, fill="#EEF2FF", outline="", tags="area")
            self.chart_canvas.create_line(
                curve_coords, fill="#6366F1", width=2.4,
                capstyle="round", joinstyle="round", tags="line"
            )
        elif len(points) == 1:
            px, py = points[0][0], points[0][1]
            base_y = pad_t + plot_h
            bar_w = 28
            self.chart_canvas.create_rectangle(
                px - bar_w / 2, py, px + bar_w / 2, base_y,
                fill="#EEF2FF", outline="#6366F1", width=1.5, tags="area"
            )
            self.chart_canvas.create_line(
                pad_l, py, pad_l + plot_w, py, fill="#E0E7FF", width=1.2,
                dash=(4, 4), tags="line"
            )

        # 6. Nodes & Peak Highlights
        actual_peak_val = max(d["gross"] for d in daily_list)
        x_step = max(1, n // 10)
        peak_accent_col = "#EF4444"  # Crimson Red Peak Accent

        for idx, (x, y, label, val, cnt, _, _) in enumerate(points):
            is_peak = (val == actual_peak_val)
            dot_col = peak_accent_col if is_peak else "#6366F1"
            r = 5 if is_peak else 3.5

            # White outer halo + center dot
            self.chart_canvas.create_oval(x - r - 2, y - r - 2, x + r + 2, y + r + 2, fill=WHITE, outline="", tags="nodes")
            self.chart_canvas.create_oval(x - r, y - r, x + r, y + r, fill=dot_col, outline=WHITE, width=2, tags="nodes")

            # X-axis label
            if idx % x_step == 0 or idx == n - 1 or is_peak:
                self.chart_canvas.create_text(x, pad_t + plot_h + 14, text=label, fill="#64748B", font=(FONT_FAMILY, 8, "bold"), tags="axis")

            # Minimalist Peak Indicator with Directional Edge Protection
            if is_peak:
                vol_str = f"Rp {val/1_000_000:,.1f}M" if val >= 1_000_000 else f"Rp {val:,.0f}"
                peak_text = f"🔥 {vol_str}"
                ico_font = ("Segoe UI Emoji", 9, "bold") if IS_WINDOWS else (FONT_FAMILY, 9, "bold")

                if x > w - pad_r - 50:
                    # Last Day / Right Edge: Number to the left, arrow pointing right ( > ) towards dot
                    self.chart_canvas.create_text(
                        x - 15, y,
                        text=peak_text, fill=peak_accent_col, font=ico_font, anchor="e", tags="peak_badge"
                    )
                    self.chart_canvas.create_polygon(
                        x - 12, y - 4,
                        x - 6, y,
                        x - 12, y + 4,
                        fill=peak_accent_col, outline="", tags="peak_badge"
                    )
                elif x < pad_l + 50:
                    # First Day / Left Edge: Number to the right, arrow pointing left ( < ) towards dot
                    self.chart_canvas.create_text(
                        x + 15, y,
                        text=peak_text, fill=peak_accent_col, font=ico_font, anchor="w", tags="peak_badge"
                    )
                    self.chart_canvas.create_polygon(
                        x + 12, y - 4,
                        x + 6, y,
                        x + 12, y + 4,
                        fill=peak_accent_col, outline="", tags="peak_badge"
                    )
                else:
                    # Middle: Centered text above centered downward arrow ( v )
                    self.chart_canvas.create_text(
                        x, y - 19,
                        text=peak_text, fill=peak_accent_col, font=ico_font, anchor="center", tags="peak_badge"
                    )
                    self.chart_canvas.create_polygon(
                        x - 4, y - 11,
                        x + 4, y - 11,
                        x, y - 6,
                        fill=peak_accent_col, outline="", tags="peak_badge"
                    )







        # Re-render active hover state if mouse is over

        if self._hover_idx is not None and 0 <= self._hover_idx < len(points):
            self._render_hover_state(self._hover_idx, pad_t, plot_h)

    def _on_canvas_motion(self, event):
        if not self._chart_points:
            return
        mx = event.x
        # Find closest point on X
        closest_idx = min(range(len(self._chart_points)), key=lambda i: abs(self._chart_points[i][0] - mx))
        if self._hover_idx != closest_idx:
            self._hover_idx = closest_idx
            pad_t = 35
            plot_h = max(40, self.chart_canvas.winfo_height() - pad_t - 30)
            self._render_hover_state(closest_idx, pad_t, plot_h)

    def _on_canvas_leave(self, event):
        self._hover_idx = None
        self.chart_canvas.delete("hover_ui")

    def _render_hover_state(self, idx: int, pad_t: int, plot_h: int):
        self.chart_canvas.delete("hover_ui")
        if not self._chart_points or idx >= len(self._chart_points):
            return

        x, y, label, val, cnt, banks_dict, full_date = self._chart_points[idx]
        w = self.chart_canvas.winfo_width()
        h = self.chart_canvas.winfo_height()

        # 1. Vertical Indicator Line
        self.chart_canvas.create_line(x, pad_t, x, pad_t + plot_h, fill="#6366F1", dash=(2, 2), width=1, tags="hover_ui")

        # 2. Highlighted Active Node
        self.chart_canvas.create_oval(x - 7, y - 7, x + 7, y + 7, fill="#6366F1", outline=WHITE, width=3, tags="hover_ui")

        # 3. Dynamic Non-Overflowing Tooltip Card
        vol_str = f"Rp {val:,.0f}".replace(",", ".")
        lines = [
            (f"📅 {full_date}", ACCENT, (FONT_FAMILY, 9, "bold")),
            (f"💵 Total: {vol_str}", SUCCESS, (FONT_FAMILY, 9, "bold")),
            (f"💳 {cnt} transactions", MUTED, (FONT_FAMILY, 8, "bold")),
        ]

        if banks_dict:
            bank_cols = {"BCA": "#0066AE", "MANDIRI": "#D97706", "BRI": "#004B87"}
            for b_name, b_vol in sorted(banks_dict.items(), key=lambda item: item[1], reverse=True):
                b_color = bank_cols.get(b_name, TEXT)
                lines.append((f"● {b_name}: Rp {b_vol/1_000_000:,.1f}M", b_color, (FONT_FAMILY, 8, "bold")))

        # Dynamic box dimension sizing based on content
        max_chars = max(len(t[0]) for t in lines)
        box_w = max(220, int(max_chars * 7.2 + 26))
        line_height = 17
        box_h = 16 + len(lines) * line_height

        # Smart positioning: Avoid edge clipping & avoid colliding with peak badge
        box_x1 = x + 14 if (x + box_w + 20 < w) else x - box_w - 14
        if box_x1 < 10:
            box_x1 = 10

        box_y1 = max(10, min(y - box_h // 2, h - box_h - 10))
        # If hovering directly at peak node, position card lower so it never obscures the peak badge
        if y < pad_t + 40:
            box_y1 = y + 16

        box_x2 = box_x1 + box_w
        box_y2 = box_y1 + box_h

        # Drop shadow & background card
        self.chart_canvas.create_rectangle(box_x1 + 2, box_y1 + 2, box_x2 + 2, box_y2 + 2, fill="#E2E8F0", outline="", tags="hover_ui")
        self.chart_canvas.create_rectangle(box_x1, box_y1, box_x2, box_y2, fill=WHITE, outline="#6366F1", width=1, tags="hover_ui")

        # Render structured text lines
        for l_idx, (text_line, text_col, text_font) in enumerate(lines):
            line_y = box_y1 + 12 + l_idx * line_height
            self.chart_canvas.create_text(box_x1 + 10, line_y, text=text_line, anchor="w", fill=text_col, font=text_font, tags="hover_ui")

