"""Custom DateInput widget with calendar picker."""
import tkinter as tk
import customtkinter as ctk
from ui.theme import (
    WHITE, BORDER_DARK, MUTED, TEXT, FONT_FAMILY,
    _EMOJI_FONT, PANEL, ACCENT, PREVIEW_BG
)

from datetime import datetime

class CTkDateInput(ctk.CTkFrame):
    def __init__(self, master, variable=None, default_date=None, date_pattern="mm/dd/yyyy", placeholder_text="MM/DD/YYYY", height=28, width=120, min_date=None, max_date=None, command=None, **kwargs):
        super().__init__(
            master, fg_color=WHITE, border_color=BORDER_DARK, border_width=1,
            corner_radius=6, height=height, width=width, **kwargs
        )
        self.pack_propagate(False)
        self._var = variable or tk.StringVar()
        self._date_pattern = date_pattern
        self._command = command
        self._min_date = min_date
        self._max_date = max_date
        
        self._cal_icon = tk.Label(
            self, text="📅", bg=WHITE, fg=MUTED,
            font=(_EMOJI_FONT, 8), cursor="hand2"
        )
        self._cal_icon.pack(side="right", padx=(0, 5), pady=2)
        
        self._entry = ctk.CTkEntry(
            self, textvariable=self._var, placeholder_text=placeholder_text,
            height=max(18, height - 4), border_width=0, fg_color="transparent", text_color=TEXT,
            font=(FONT_FAMILY, 10, "bold")
        )
        self._entry.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=2)


        
        self._cal_icon.bind("<Button-1>", lambda e: self.open_calendar())
        self._entry.bind("<Button-1>", lambda e: self.open_calendar())
        self.bind("<Button-1>", lambda e: self.open_calendar())
        
        if default_date:
            self.set_date(default_date)
            
    def get(self):
        return self._var.get().strip()
        
    def set_date(self, d):
        if hasattr(d, "strftime"):
            s = d.strftime("%Y-%m-%d") if self._date_pattern == "yyyy-mm-dd" else d.strftime("%m/%d/%Y")
        else:
            s = str(d)
        self._var.set(s)
        self._entry.delete(0, "end")
        self._entry.insert(0, s)
        if self._command:
            try:
                self._command(s)
            except Exception:
                pass
            
    def open_calendar(self):
        try:
            import tkcalendar
            top = ctk.CTkToplevel(self)
            top.withdraw()
            top.title("Select Date")
            top.configure(fg_color=PANEL)
            top.transient(self.winfo_toplevel())
            
            x = self.winfo_rootx()
            y = self.winfo_rooty() + self.winfo_height() + 4
            top.geometry(f"260x220+{x}+{y}")

            cal_kwargs = {}
            if self._min_date:
                m_val = self._min_date() if callable(self._min_date) else self._min_date
                if m_val:
                    if isinstance(m_val, str):
                        try:
                            m_val = datetime.strptime(m_val, "%Y-%m-%d" if self._date_pattern == "yyyy-mm-dd" else "%m/%d/%Y").date()
                        except Exception:
                            m_val = None
                    elif hasattr(m_val, "date") and callable(getattr(m_val, "date")):
                        m_val = m_val.date()
                    if m_val:
                        cal_kwargs["mindate"] = m_val

            if self._max_date:
                m_val = self._max_date() if callable(self._max_date) else self._max_date
                if m_val:
                    if isinstance(m_val, str):
                        try:
                            m_val = datetime.strptime(m_val, "%Y-%m-%d" if self._date_pattern == "yyyy-mm-dd" else "%m/%d/%Y").date()
                        except Exception:
                            m_val = None
                    elif hasattr(m_val, "date") and callable(getattr(m_val, "date")):
                        m_val = m_val.date()
                    if m_val:
                        cal_kwargs["maxdate"] = m_val

            cal = tkcalendar.Calendar(
                top, selectmode="day", date_pattern=self._date_pattern,
                background=ACCENT, foreground=WHITE, headersbackground=PREVIEW_BG,
                headersforeground=TEXT, selectbackground=ACCENT, selectforeground=WHITE,
                normalbackground=WHITE, normalforeground=TEXT,
                **cal_kwargs
            )
            cal.pack(fill="both", expand=True, padx=8, pady=8)

            
            def _select():
                self.set_date(cal.get_date())
                top.destroy()
                
            cal.bind("<<CalendarSelected>>", lambda e: _select())
            top.update_idletasks()
            top.deiconify()
            top.grab_set()
        except Exception:
            pass

