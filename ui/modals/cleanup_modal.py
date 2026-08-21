"""Data Cleanup and Workspace Recap Archive confirmation modal."""
import customtkinter as ctk
from pathlib import Path
from ui.theme import (
    BG, PANEL, PREVIEW_BG, BORDER, BORDER_DARK, TEXT, MUTED,
    WHITE, ACCENT, ERROR, FONT_FAMILY, FONT_MONO
)

def open_cleanup_modal(parent, files_to_move: list[Path], target_recap: Path, on_confirm):
    """Show modal confirming cleanup of files to recap directory."""
    if not files_to_move:
        return None

    dlg = ctk.CTkToplevel(parent)
    dlg.withdraw()
    dlg.title("Confirm Data Cleanup")
    dlg.resizable(False, False)
    dlg.transient(parent)

    sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
    cx, cy = max(0, int(sw / 2 - 540 / 2)), max(0, int(sh / 2 - 370 / 2))
    dlg.geometry(f"540x370+{cx}+{cy}")
    dlg.configure(fg_color=BG)

    content = ctk.CTkFrame(dlg, fg_color=PANEL, corner_radius=10, border_color=BORDER, border_width=1)
    content.pack(fill="both", expand=True, padx=16, pady=16)

    hdr = ctk.CTkFrame(content, fg_color="transparent")
    hdr.pack(fill="x", padx=16, pady=(16, 8))
    ctk.CTkLabel(hdr, text=f"This will relocate {len(files_to_move)} file(s) out of active input, mutation & output folders.", font=(FONT_FAMILY, 10, "bold"), text_color=MUTED).pack(anchor="w", pady=(2, 0))

    info_box = ctk.CTkFrame(content, fg_color=PREVIEW_BG, corner_radius=8, border_color=BORDER, border_width=1)
    info_box.pack(fill="x", padx=16, pady=12)

    ctk.CTkLabel(info_box, text="🗄️ Archive Destination Folder:", font=(FONT_FAMILY, 10, "bold"), text_color=TEXT).pack(anchor="w", padx=12, pady=(10, 2))
    ctk.CTkLabel(info_box, text=str(target_recap), font=(FONT_MONO, 10, "bold"), text_color=ACCENT, wraplength=460, justify="left").pack(anchor="w", padx=12, pady=(0, 10))

    ctk.CTkLabel(content, text="Cleaned files will be stored safely in Recap storage and can be accessed via Quick Access.", font=(FONT_FAMILY, 10, "bold"), text_color=MUTED, wraplength=460, justify="left").pack(anchor="w", padx=16)

    btn_frame = ctk.CTkFrame(content, fg_color="transparent")
    btn_frame.pack(fill="x", side="bottom", padx=16, pady=(12, 16))

    def _do_clean():
        dlg.destroy()
        if on_confirm:
            on_confirm()

    ctk.CTkButton(
        btn_frame, text="Cancel", height=36, width=100,
        fg_color=WHITE, hover_color=PREVIEW_BG, border_color=BORDER_DARK, border_width=1,
        text_color=TEXT, font=(FONT_FAMILY, 11, "bold"), command=dlg.destroy
    ).pack(side="right", padx=(8, 0))

    ctk.CTkButton(
        btn_frame, text="Confirm Clean", height=36, width=130,
        fg_color=ERROR, hover_color="#B91C1C", text_color=WHITE,
        font=(FONT_FAMILY, 11, "bold"), command=_do_clean
    ).pack(side="right")

    dlg.update_idletasks()
    dlg.deiconify()
    dlg.grab_set()
    return dlg
