"""
Terminal Log Output Console View with syntax color-tagging and status bar.
"""
import tkinter as tk
from datetime import datetime
import customtkinter as ctk

from ui.theme import (
    PANEL, PREVIEW_BG, BORDER, BORDER_DARK, ACCENT, SUCCESS, ERROR, WARN, TEXT, MUTED, WHITE,
    FONT_FAMILY, FONT_MONO, IS_WINDOWS
)


class LogConsoleView(ctk.CTkFrame):
    def __init__(self, master, on_clear=None, **kwargs):
        super().__init__(master, fg_color=PANEL, corner_radius=10, border_color=BORDER, border_width=1, **kwargs)
        self.on_clear = on_clear
        self._build_ui()

    def _build_ui(self):
        # 1. Console Header Toolbar
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=14, pady=(6, 4))

        ico_font = ("Segoe UI Emoji", 10) if IS_WINDOWS else (FONT_FAMILY, 10)
        lbl_term_ico = tk.Label(hdr, text="⚡", font=ico_font, bg=PANEL, fg=ACCENT)
        lbl_term_ico.pack(side="left", padx=(0, 4))

        ctk.CTkLabel(
            hdr, text="TERMINAL LOG OUTPUT",
            font=(FONT_FAMILY, 10, "bold"), text_color=MUTED,
            fg_color="transparent"
        ).pack(side="left")

        self.btn_clear = ctk.CTkButton(
            hdr, text="Clear Log", width=65, height=22,
            fg_color="transparent", hover_color=PREVIEW_BG,
            border_color=BORDER_DARK, border_width=1,
            text_color=MUTED, font=(FONT_FAMILY, 9, "bold"),
            corner_radius=5, command=self.clear
        )
        self.btn_clear.pack(side="right")

        # 2. Text Box with Scrollbar
        log_frame = ctk.CTkFrame(self, fg_color=PREVIEW_BG, corner_radius=6, border_color=BORDER, border_width=1)
        log_frame.pack(fill="both", expand=True, padx=12, pady=(0, 6))


        self.log_text = tk.Text(
            log_frame,
            bg=PREVIEW_BG,
            fg=TEXT,
            font=(FONT_MONO, 10),
            relief="flat",
            wrap="word",
            state="disabled",
            insertbackground=ACCENT,
            selectbackground=BORDER_DARK,
            padx=12,
            pady=10,
            bd=0,
            highlightthickness=0
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar = ctk.CTkScrollbar(
            log_frame, orientation="vertical", command=self.log_text.yview,
            button_color=BORDER_DARK, button_hover_color=MUTED
        )
        scrollbar.pack(side="right", fill="y", padx=(0, 2), pady=2)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        # Configure Log Tag Styles
        self.log_text.tag_config("head", foreground=ACCENT, font=(FONT_MONO, 10, "bold"))
        self.log_text.tag_config("ok",   foreground=SUCCESS, font=(FONT_MONO, 10, "bold"))
        self.log_text.tag_config("warn", foreground=WARN, font=(FONT_MONO, 10, "bold"))
        self.log_text.tag_config("err",  foreground=ERROR, font=(FONT_MONO, 10, "bold"))
        self.log_text.tag_config("dim",  foreground=MUTED, font=(FONT_MONO, 10))

        # 3. Status Bar Footer
        status_bar = ctk.CTkFrame(self, fg_color=PANEL, height=28)
        status_bar.pack(fill="x", side="bottom", padx=14, pady=(0, 6))

        self.status_dot = ctk.CTkLabel(
            status_bar, text="●", font=(FONT_FAMILY, 12, "bold"),
            text_color=SUCCESS, fg_color="transparent", width=14
        )
        self.status_dot.pack(side="left")

        self.status_var = tk.StringVar(value="Ready")
        self.lbl_status = ctk.CTkLabel(
            status_bar, textvariable=self.status_var,
            font=(FONT_FAMILY, 10, "bold"), text_color=TEXT,
            fg_color="transparent", anchor="w"
        )
        self.lbl_status.pack(side="left", padx=(4, 0))

    def write(self, text: str, tag: str = ""):
        """Write formatted line to terminal log."""
        self.log_text.config(state="normal")
        self.log_text.insert("end", text, tag)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def clear(self):
        """Clear all terminal text."""
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state="disabled")
        if callable(self.on_clear):
            self.on_clear()

    def set_status(self, text: str, color: str):
        """Update bottom status badge text and indicator color."""
        self.status_var.set(text)
        self.status_dot.configure(text_color=color)
