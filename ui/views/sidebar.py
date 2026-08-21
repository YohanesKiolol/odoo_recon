"""
Sidebar View Component for Bank Reconciliation Studio.
Includes Workspace / Company selector, Credentials, Bank targets, Date range, Quick Access & Action CTA stack.
"""
import tkinter as tk
import customtkinter as ctk

from ui.theme import (
    SIDEBAR_BG, PANEL, PREVIEW_BG, BORDER, BORDER_DARK, ACCENT, ACCENT_DARK,
    SUCCESS, ERROR, WARN, TEXT, MUTED, WHITE, FONT_FAMILY, IS_WINDOWS
)
from ui.widgets import CTkDateInput, create_custom_dropdown


class SidebarView(ctk.CTkFrame):
    def __init__(
        self,
        master,
        on_run=None,
        on_stop=None,
        on_manual_match=None,
        on_sync_cloud=None,
        on_journal=None,
        on_company_change=None,
        on_source_change=None,
        quick_access_handlers=None,
        **kwargs
    ):
        super().__init__(master, width=280, fg_color=SIDEBAR_BG, corner_radius=0, border_color=BORDER, border_width=1, **kwargs)
        self.pack_propagate(False)

        self.on_run = on_run
        self.on_stop = on_stop
        self.on_manual_match = on_manual_match
        self.on_sync_cloud = on_sync_cloud
        self.on_journal = on_journal
        self.on_company_change = on_company_change
        self.on_source_change = on_source_change
        self.qa_handlers = quick_access_handlers or {}

        self._bank_vars = {}
        self._card_refs = {}
        self.email_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.date_from_var = tk.StringVar()
        self.date_to_var = tk.StringVar()
        self.offline_var = tk.BooleanVar(value=False)
        self.recon_source_var = tk.StringVar(value="local")

        self._build_ui()

    def _build_ui(self):
        # Top Logo / Title
        brand_frame = ctk.CTkFrame(self, fg_color="transparent")
        brand_frame.pack(fill="x", padx=16, pady=(16, 10))

        lbl_logo = ctk.CTkLabel(
            brand_frame, text="⚡ ReconStudio",
            font=(FONT_FAMILY, 15, "bold"), text_color=TEXT
        )
        lbl_logo.pack(side="left")

        lbl_ver = ctk.CTkLabel(
            brand_frame, text="v2.4",
            font=(FONT_FAMILY, 9, "bold"), text_color=MUTED,
            fg_color=PANEL, corner_radius=4, padx=6, pady=2
        )
        lbl_ver.pack(side="right")

        ctk.CTkFrame(self, height=1, fg_color=BORDER).pack(fill="x")

        # Scrollable Control Area
        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        # ── 1. Workspace / Company Section ──
        sec_workspace = ctk.CTkFrame(scroll, fg_color=PANEL, corner_radius=8, border_color=BORDER, border_width=1)
        sec_workspace.pack(fill="x", padx=6, pady=4)

        ctk.CTkLabel(
            sec_workspace, text="W O R K S P A C E   &   U S E R", font=(FONT_FAMILY, 9, "bold"),
            text_color=MUTED, fg_color="transparent"
        ).pack(anchor="w", padx=10, pady=(8, 4))

        ctk.CTkLabel(
            sec_workspace, text="Company / Outlet", font=(FONT_FAMILY, 9, "bold"),
            text_color=TEXT, fg_color="transparent"
        ).pack(anchor="w", padx=10, pady=(2, 2))

        import config
        avail_companies = config.get_available_companies()
        curr_name = config.COMPANY_NAME

        self.btn_comp = create_custom_dropdown(
            sec_workspace,
            label=f"🏢  {curr_name}",
            options=[(f"  🏢  {c_n}  ", lambda k=c_k: self._select_company(k)) for c_k, c_n in avail_companies.items()],
            height=34
        )
        self.btn_comp.pack(fill="x", padx=10, pady=(0, 8))

        # Credentials Sub-block
        ctk.CTkLabel(
            sec_workspace, text="Odoo Profile", font=(FONT_FAMILY, 9, "bold"),
            text_color=TEXT, fg_color="transparent"
        ).pack(anchor="w", padx=10, pady=(2, 2))

        self.cred_container = ctk.CTkFrame(sec_workspace, fg_color="transparent")
        self.cred_container.pack(fill="x", padx=10, pady=(0, 10))
        self.rebuild_credentials_ui()

        # ── 2. Bank Target Section ──
        self.sec_bank = ctk.CTkFrame(scroll, fg_color=PANEL, corner_radius=8, border_color=BORDER, border_width=1)
        self.sec_bank.pack(fill="x", padx=6, pady=4)

        ctk.CTkLabel(
            self.sec_bank, text="B A N K   T A R G E T", font=(FONT_FAMILY, 9, "bold"),
            text_color=MUTED, fg_color="transparent"
        ).pack(anchor="w", padx=10, pady=(8, 6))

        self.bank_grid = ctk.CTkFrame(self.sec_bank, fg_color="transparent")
        self.bank_grid.pack(fill="x", padx=8, pady=(0, 10))
        self.rebuild_bank_cards()

        # ── 3. Date Range Section ──
        sec_date = ctk.CTkFrame(scroll, fg_color=PANEL, corner_radius=8, border_color=BORDER, border_width=1)
        sec_date.pack(fill="x", padx=6, pady=4)

        ctk.CTkLabel(
            sec_date, text="D A T E   R A N G E", font=(FONT_FAMILY, 9, "bold"),
            text_color=MUTED, fg_color="transparent"
        ).pack(anchor="w", padx=10, pady=(8, 4))

        date_grid = ctk.CTkFrame(sec_date, fg_color="transparent")
        date_grid.pack(fill="x", padx=8, pady=(0, 8))
        date_grid.columnconfigure(0, weight=1, uniform="date_col")
        date_grid.columnconfigure(1, weight=1, uniform="date_col")

        ctk.CTkLabel(date_grid, text="From", font=(FONT_FAMILY, 9, "bold"), text_color=MUTED, fg_color="transparent").grid(row=0, column=0, sticky="w", padx=2)
        ctk.CTkLabel(date_grid, text="To", font=(FONT_FAMILY, 9, "bold"), text_color=MUTED, fg_color="transparent").grid(row=0, column=1, sticky="w", padx=2)

        self.date_from_widget = CTkDateInput(date_grid, variable=self.date_from_var, default_date=None)
        self.date_from_widget.grid(row=1, column=0, sticky="ew", padx=(0, 3))

        self.date_to_widget = CTkDateInput(date_grid, variable=self.date_to_var, default_date=None)
        self.date_to_widget.grid(row=1, column=1, sticky="ew", padx=(3, 0))

        # ── 4. Quick Access Section ──
        sec_links = ctk.CTkFrame(scroll, fg_color=PANEL, corner_radius=8, border_color=BORDER, border_width=1)
        sec_links.pack(fill="x", padx=6, pady=4)

        ctk.CTkLabel(
            sec_links, text="Q U I C K   A C C E S S", font=(FONT_FAMILY, 9, "bold"),
            text_color=MUTED, fg_color="transparent"
        ).pack(anchor="w", padx=10, pady=(6, 2))

        def _link_btn(parent, icon, label_text, key):
            row = ctk.CTkFrame(parent, fg_color="transparent", height=28)
            row.pack(fill="x", padx=4, pady=1)
            row.pack_propagate(False)

            cmd = self.qa_handlers.get(key, lambda: None)
            ico_font = ("Segoe UI Emoji", 10) if IS_WINDOWS else (FONT_FAMILY, 10)
            lbl_ico = tk.Label(row, text=icon, font=ico_font, bg=PANEL, fg=MUTED, width=3, anchor="center")
            lbl_ico.pack(side="left", padx=(4, 2))
            lbl_ico.bind("<Button-1>", lambda e: cmd())

            btn = ctk.CTkButton(
                row, text=label_text, height=28,
                fg_color="transparent", hover_color=PREVIEW_BG,
                text_color=ACCENT, font=(FONT_FAMILY, 10, "bold"),
                anchor="w", command=cmd
            )
            btn.pack(side="left", fill="both", expand=True)
            return btn

        _link_btn(sec_links, "📊", "Open Result", "open_output")
        _link_btn(sec_links, "📄", "Summary PDF", "export_pdf")
        _link_btn(sec_links, "📂", "Open Merchant", "open_input")
        _link_btn(sec_links, "📁", "Open Mutation", "open_mutation")
        _link_btn(sec_links, "💳", "Open Payment", "open_odoo_file")
        _link_btn(sec_links, "📑", "Open Journal Entries", "open_journal_file")
        _link_btn(sec_links, "🗄", "Open Recap", "open_recap")

        # ── 5. Primary CTA Action Stack ──
        ctk.CTkFrame(self, height=1, fg_color=BORDER).pack(fill="x")
        _cta = ctk.CTkFrame(self, fg_color="transparent")
        _cta.pack(fill="x", padx=12, pady=12)

        ctk.CTkCheckBox(
            _cta, text="Offline Mode (Skip Downloader)",
            variable=self.offline_var,
            font=(FONT_FAMILY, 9, "bold"),
            fg_color=ACCENT, hover_color=ACCENT_DARK,
            text_color=TEXT, checkbox_width=16, checkbox_height=16
        ).pack(anchor="w", padx=2, pady=(0, 6))

        # Data Source Selector: Local vs Cloud
        src_box = ctk.CTkFrame(_cta, fg_color="transparent")
        src_box.pack(fill="x", pady=(0, 6))

        ctk.CTkLabel(src_box, text="Source:", font=(FONT_FAMILY, 9, "bold"), text_color=MUTED).pack(side="left", padx=(2, 6))
        
        self.btn_src_local = ctk.CTkButton(
            src_box, text="📁 Local", height=24, width=64,
            fg_color=ACCENT, text_color=WHITE, font=(FONT_FAMILY, 9, "bold"), corner_radius=5,
            command=lambda: self.set_recon_source("local")
        )
        self.btn_src_local.pack(side="left", padx=(0, 4))

        self.btn_src_cloud = ctk.CTkButton(
            src_box, text="☁️ Cloud", height=24, width=64,
            fg_color=WHITE, text_color=TEXT, border_color=BORDER_DARK, border_width=1,
            font=(FONT_FAMILY, 9, "bold"), corner_radius=5,
            command=lambda: self.set_recon_source("cloud")
        )
        self.btn_src_cloud.pack(side="left")

        self.run_btn = ctk.CTkButton(
            _cta, text="⚡  Run Reconciliation",
            height=40, fg_color=WHITE, hover_color="#F8FAFC",
            border_color=BORDER_DARK, border_width=1,
            text_color=TEXT, font=(FONT_FAMILY, 11, "bold"),
            corner_radius=8, command=self.on_run
        )
        self.run_btn.pack(fill="x", pady=(0, 6))

        self.match_btn = ctk.CTkButton(
            _cta, text="🧩  Manual Match",
            height=40, fg_color=WHITE, hover_color="#F8FAFC",
            border_color=BORDER_DARK, border_width=1,
            text_color=ACCENT, font=(FONT_FAMILY, 11, "bold"),
            corner_radius=8, command=self.on_manual_match
        )
        self.match_btn.pack(fill="x", pady=(0, 6))

        self.sync_cloud_btn = ctk.CTkButton(
            _cta, text="☁️  Sync Sales Portal",
            height=40, fg_color=WHITE, hover_color="#F8FAFC",
            border_color=BORDER_DARK, border_width=1,
            text_color=TEXT, font=(FONT_FAMILY, 11, "bold"),
            corner_radius=8, command=self.on_sync_cloud
        )
        self.sync_cloud_btn.pack(fill="x", pady=(0, 6))

        self.journal_btn = ctk.CTkButton(
            _cta, text="📋  Generate Journal",
            height=42, fg_color=ACCENT, hover_color=ACCENT_DARK,
            text_color=WHITE, font=(FONT_FAMILY, 11, "bold"),
            corner_radius=8, command=self.on_journal
        )
        self.journal_btn.pack(fill="x")

        self.stop_btn = ctk.CTkButton(
            _cta, text="⏹  Stop Process",
            height=38, fg_color=ERROR, hover_color="#B91C1C",
            text_color=WHITE, font=(FONT_FAMILY, 11, "bold"),
            corner_radius=8, command=self.on_stop
        )

    def _select_company(self, ckey):
        import config
        config.load_company_env(ckey)
        new_name = config.COMPANY_NAME
        self.btn_comp.configure(text=f"🏢  {new_name}")
        self.rebuild_credentials_ui()
        self.rebuild_bank_cards()
        if callable(self.on_company_change):
            self.on_company_change(ckey, new_name)

    def set_recon_source(self, mode: str):
        self.recon_source_var.set(mode)
        if mode == "cloud":
            self.btn_src_cloud.configure(fg_color=ACCENT, text_color=WHITE, border_width=0)
            self.btn_src_local.configure(fg_color=WHITE, text_color=TEXT, border_color=BORDER_DARK, border_width=1)
        else:
            self.btn_src_local.configure(fg_color=ACCENT, text_color=WHITE, border_width=0)
            self.btn_src_cloud.configure(fg_color=WHITE, text_color=TEXT, border_color=BORDER_DARK, border_width=1)
        if callable(self.on_source_change):
            self.on_source_change(mode)

    def rebuild_credentials_ui(self):
        for child in self.cred_container.winfo_children():
            child.destroy()

        import config
        predefined = config.get_predefined_accounts()
        if predefined:
            preset_names = list(predefined.keys())
            first_key = preset_names[0]

            def _select_acc(name):
                if name in predefined:
                    acc = predefined[name]
                    self.email_var.set(acc["username"])
                    key_val = acc.get("api_key") or acc.get("password", "")
                    self.password_var.set(key_val)
                    self.btn_preset.configure(text=f"👤  {name}")
                    try:
                        import odoo_inspector
                        odoo_inspector.set_active_credentials(acc["username"], key_val)
                    except Exception:
                        pass

            self.btn_preset = create_custom_dropdown(
                self.cred_container,
                label=f"👤  {first_key}",
                options=[(f"  👤  {n}  ", lambda name=n: _select_acc(name)) for n in preset_names],
                height=34
            )
            self.btn_preset.pack(fill="x")
            _select_acc(first_key)
        else:
            ctk.CTkEntry(
                self.cred_container, textvariable=self.email_var,
                placeholder_text="odoo@example.com",
                height=32, corner_radius=6, border_color=BORDER_DARK, fg_color=WHITE, text_color=TEXT,
                font=(FONT_FAMILY, 10, "bold")
            ).pack(fill="x")

    def rebuild_bank_cards(self):
        for child in self.bank_grid.winfo_children():
            child.destroy()
        self._card_refs.clear()

        import config
        configured_banks = list(getattr(config, "CONFIGURED_BANKS", ["BCA", "Mandiri", "BRI"]))
        bank_list = ["All"] + configured_banks if len(configured_banks) > 1 else configured_banks
        subtitles = getattr(config, "BANK_SUBTITLES", {})
        _strip_col = {"All": ACCENT, "BCA": "#0066AE", "Mandiri": "#F0A500", "BRI": "#004B87"}

        self._bank_vars.clear()
        for b in bank_list:
            self._bank_vars[b] = tk.BooleanVar(value=(b == "All" or len(bank_list) == 1))

        def _on_bank_toggle(name):
            if name == "All" and self._bank_vars.get("All") and self._bank_vars["All"].get():
                for b in configured_banks:
                    if b in self._bank_vars:
                        self._bank_vars[b].set(False)
            elif name in configured_banks and self._bank_vars.get(name) and self._bank_vars[name].get():
                if "All" in self._bank_vars:
                    self._bank_vars["All"].set(False)
            self._refresh_bank_cards()

        def _refresh_bank_cards():
            for bname, refs in self._card_refs.items():
                outer, strip, name_lbl, sub_lbl, check_lbl = refs
                sel = self._bank_vars[bname].get() if bname in self._bank_vars else False
                strip_c = _strip_col.get(bname, ACCENT)
                outer.configure(
                    fg_color=WHITE if sel else PREVIEW_BG,
                    border_color=strip_c if sel else BORDER,
                    border_width=2 if sel else 1,
                )
                strip.configure(fg_color=strip_c if sel else BORDER_DARK)
                name_lbl.configure(text_color=strip_c if sel else TEXT)
                sub_lbl.configure(text_color=MUTED)
                check_lbl.configure(text="✓" if sel else "", text_color=strip_c)

        self._refresh_bank_cards = _refresh_bank_cards

        self.bank_grid.columnconfigure(0, weight=1, uniform="bank_col")
        self.bank_grid.columnconfigure(1, weight=1 if len(bank_list) > 1 else 0, uniform="bank_col")

        for idx, bname in enumerate(bank_list):
            grow = idx // 2
            gcol = idx % 2 if len(bank_list) > 1 else 0
            sel = self._bank_vars[bname].get()
            strip_c = _strip_col.get(bname, ACCENT)

            def _make_cmd(n=bname):
                def _cmd():
                    if n in self._bank_vars:
                        self._bank_vars[n].set(not self._bank_vars[n].get())
                        _on_bank_toggle(n)
                return _cmd

            outer = ctk.CTkFrame(
                self.bank_grid, width=120, height=48, corner_radius=7,
                fg_color=WHITE if sel else PREVIEW_BG,
                border_color=strip_c if sel else BORDER,
                border_width=2 if sel else 1,
                cursor="hand2",
            )
            if len(bank_list) == 1:
                outer.grid(row=grow, column=0, columnspan=2, padx=2, pady=3, sticky="nsew")
            else:
                outer.grid(row=grow, column=gcol, padx=2, pady=3, sticky="nsew")
            outer.pack_propagate(False)
            outer.bind("<Button-1>", lambda e, n=bname: _make_cmd(n)())

            strip = ctk.CTkFrame(outer, width=4, corner_radius=0, fg_color=strip_c if sel else BORDER_DARK)
            strip.pack(side="left", fill="y")
            strip.pack_propagate(False)
            strip.bind("<Button-1>", lambda e, n=bname: _make_cmd(n)())

            txt = ctk.CTkFrame(outer, fg_color="transparent")
            txt.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=3)
            txt.bind("<Button-1>", lambda e, n=bname: _make_cmd(n)())

            name_lbl = ctk.CTkLabel(
                txt, text=bname if bname != "All" else "All Banks",
                font=(FONT_FAMILY, 10, "bold"), text_color=strip_c if sel else TEXT,
                fg_color="transparent", anchor="w",
            )
            name_lbl.pack(anchor="w")
            name_lbl.bind("<Button-1>", lambda e, n=bname: _make_cmd(n)())

            sub_lbl = ctk.CTkLabel(
                txt, text=subtitles.get(bname, ""), font=(FONT_FAMILY, 8), text_color=MUTED,
                fg_color="transparent", anchor="w",
            )
            sub_lbl.pack(anchor="w")
            sub_lbl.bind("<Button-1>", lambda e, n=bname: _make_cmd(n)())

            check_lbl = ctk.CTkLabel(
                outer, text="✓" if sel else "", font=(FONT_FAMILY, 11, "bold"),
                text_color=strip_c, fg_color="transparent", width=14,
            )
            check_lbl.pack(side="right", padx=(0, 4))
            check_lbl.bind("<Button-1>", lambda e, n=bname: _make_cmd(n)())

            self._card_refs[bname] = (outer, strip, name_lbl, sub_lbl, check_lbl)

        _refresh_bank_cards()

    def get_selected_banks(self) -> list[str]:
        """Return selected bank keys (e.g. ['bca', 'mandiri', 'bri'])."""
        if self._bank_vars.get("All") and self._bank_vars["All"].get():
            return ["bca", "mandiri", "bri"]
        sel = [b.lower() for b in ["BCA", "Mandiri", "BRI"] if self._bank_vars.get(b) and self._bank_vars[b].get()]
        return sel or ["bca", "mandiri", "bri"]
