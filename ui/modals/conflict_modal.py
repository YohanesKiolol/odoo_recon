"""Upload file conflict resolution modal dialog."""
import tkinter as tk
import customtkinter as ctk
from ui.theme import (
    BG, PANEL, BORDER, BORDER_DARK, PREVIEW_BG, TEXT, MUTED,
    WHITE, ACCENT, ACCENT_HOVER, ERROR, WARN, FONT_FAMILY, FONT_MONO
)

def open_conflict_modal(parent, conflicts: list) -> str:
    """Show modal to resolve destination file conflicts and return action chosen."""
    if not conflicts:
        return "replace"

    names = "\n".join(f"  • {p.name}" for p, _, _ in conflicts[:8])
    if len(conflicts) > 8:
        names += f"\n  ... and {len(conflicts) - 8} more"

    dlg = ctk.CTkToplevel(parent)
    dlg.withdraw()
    dlg.title("File Conflict")
    dlg.resizable(False, False)
    dlg.transient(parent)
    dlg.configure(fg_color=BG)

    sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
    cx, cy = max(0, sw // 2 - 240), max(0, sh // 2 - 140)
    dlg.geometry(f"480x280+{cx}+{cy}")

    choice = tk.StringVar(value="")
    panel = ctk.CTkFrame(dlg, fg_color=PANEL, corner_radius=10, border_color=BORDER, border_width=1)
    panel.pack(fill="both", expand=True, padx=14, pady=14)
    ctk.CTkLabel(panel, text=f"⚠️  {len(conflicts)} file(s) already exist in destination:", font=(FONT_FAMILY, 12, "bold"), text_color=WARN).pack(anchor="w", padx=14, pady=(14, 4))
    ctk.CTkLabel(panel, text=names, font=(FONT_MONO, 10), text_color=MUTED, justify="left").pack(anchor="w", padx=14)

    btn_frame = ctk.CTkFrame(panel, fg_color="transparent")
    btn_frame.pack(fill="x", side="bottom", padx=14, pady=14)

    def _pick(v):
        choice.set(v)
        dlg.destroy()

    ctk.CTkButton(btn_frame, text="Replace All", width=120, height=34, fg_color=ERROR, hover_color="#B91C1C", text_color=WHITE, font=(FONT_FAMILY, 11, "bold"), command=lambda: _pick("replace")).pack(side="left", padx=(0, 8))
    ctk.CTkButton(btn_frame, text="Keep Both", width=120, height=34, fg_color=ACCENT, hover_color=ACCENT_HOVER, text_color=WHITE, font=(FONT_FAMILY, 11, "bold"), command=lambda: _pick("keep")).pack(side="left", padx=(0, 8))
    ctk.CTkButton(btn_frame, text="Skip Conflicts", width=120, height=34, fg_color=WHITE, hover_color=PREVIEW_BG, border_color=BORDER_DARK, border_width=1, text_color=TEXT, font=(FONT_FAMILY, 11, "bold"), command=lambda: _pick("skip")).pack(side="left")

    dlg.update_idletasks()
    dlg.deiconify()
    dlg.grab_set()
    dlg.wait_window()
    return choice.get() or "skip"
