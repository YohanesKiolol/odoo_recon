"""Custom bordered dropdown menu component matching the Workspace design."""
import tkinter as tk
import customtkinter as ctk
from ui.theme import (
    WHITE, BORDER_DARK, MUTED, TEXT, ACCENT, PREVIEW_BG, FONT_FAMILY
)


def create_custom_dropdown(
    parent,
    variable: tk.StringVar | None = None,
    values: list[str] | None = None,
    label: str | None = None,
    options: list[tuple[str, any]] | None = None,
    width: int = 140,
    height: int = 30,
    on_change = None
) -> ctk.CTkButton:
    """
    Create a unified dropdown button matching the Workspace design system.
    Supports either variable+values or label+options list.
    """
    if options:
        display_text = label or (options[0][0] if options else "")
    else:
        current_val = variable.get() if variable else ""
        if not current_val and values:
            current_val = values[0]
            if variable:
                variable.set(values[0])
        display_text = label or f"  {current_val}"

    btn = ctk.CTkButton(
        parent,
        text=display_text,
        anchor="w",
        height=height,
        width=width,
        corner_radius=6,
        border_width=1,
        border_color=BORDER_DARK,
        fg_color=WHITE,
        hover_color=PREVIEW_BG,
        text_color=TEXT,
        font=(FONT_FAMILY, 10, "bold"),
    )

    lbl_chev = ctk.CTkLabel(
        btn, text="▾", text_color=MUTED,
        font=(FONT_FAMILY, 10, "bold"), fg_color="transparent"
    )
    lbl_chev.place(relx=1.0, rely=0.5, anchor="e", x=-8)

    menu = tk.Menu(
        btn, tearoff=0, bg=WHITE, fg=TEXT,
        activebackground=PREVIEW_BG, activeforeground=ACCENT,
        font=(FONT_FAMILY, 10, "bold"), bd=1, relief="solid"
    )

    if options:
        for opt_label, opt_cmd in options:
            menu.add_command(
                label=opt_label,
                command=opt_cmd
            )
    elif values:
        def _select_item(item_val):
            if variable:
                variable.set(item_val)
            btn.configure(text=f"  {item_val}")
            if on_change:
                on_change(item_val)

        for val in values:
            menu.add_command(
                label=f"  {val}  ",
                command=lambda v=val: _select_item(v)
            )

    def _open_menu(event=None):
        try:
            x = btn.winfo_rootx()
            y = btn.winfo_rooty() + btn.winfo_height() + 2
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    btn.configure(command=_open_menu)
    lbl_chev.bind("<Button-1>", _open_menu)

    return btn
