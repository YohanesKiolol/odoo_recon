"""Design system, color palette, typography and platform constants for Bank Reconciliation Studio."""
import os
import sys
from pathlib import Path
import customtkinter as ctk

# Base Directory
BASE_DIR = (
    Path(sys._MEIPASS)
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent.parent
)

# Platform Flags
IS_WINDOWS = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"

# Typography
FONT_FAMILY = "Space Grotesk"
FONT_BODY   = "IBM Plex Sans"
FONT_MONO   = "Consolas" if IS_WINDOWS else "Menlo"
_EMOJI_FONT = "Segoe UI Emoji" if IS_WINDOWS else FONT_FAMILY

# Theme Palette (Modern High-Contrast Clean Light)
BG          = "#F4F5F8"  # Cool neutral app background
PANEL       = "#FFFFFF"  # Pure white card background
SIDEBAR_BG  = "#FAFAFC"  # Soft off-white sidebar background
PREVIEW_BG  = "#F8FAFC"  # Table/Sub-card background
BORDER      = "#E2E8F0"  # Crisp subtle border color
BORDER_DARK = "#CBD5E1"  # Stronger border color for inputs

ACCENT      = "#6D28D9"  # Odoo Deep Violet
ACCENT_DARK = "#5B21B6"  # Hover Violet
ACCENT_HOVER= "#5B21B6"
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

BANK_BADGE_COLS = {
    "BCA":     ("#1E40AF", "#DBEAFE"),
    "MANDIRI": ("#92400E", "#FEF3C7"),
    "BRI":     ("#047857", "#D1FAE5"),
    "OTHER":   ("#4B5563", "#F3F4F6"),
}

TYPE_BADGE_INFO = {
    "bank_only":         ("🏦 Bank Only", "#4338CA", "#EEF2FF"),
    "odoo_only":         ("📦 Odoo Only", "#B45309", "#FEF3C7"),
    "unreconciled_odoo": ("⚠️ Unreconciled", "#BE123C", "#FFE4E6"),
}

def init_fonts():
    """Load bundled custom fonts."""
    fonts_dir = (
        Path(getattr(sys, "_MEIPASS", BASE_DIR)) / "assets" / "fonts"
        if getattr(sys, "frozen", False)
        else BASE_DIR / "assets" / "fonts"
    )
    if fonts_dir.exists():
        for f_file in fonts_dir.glob("*.ttf"):
            try:
                ctk.FontManager.load_font(str(f_file))
            except Exception:
                pass


