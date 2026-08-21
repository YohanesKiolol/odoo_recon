"""Cross-platform window utilities, modal positioning and file openers."""
import sys
import os
import subprocess
from ui.theme import IS_WINDOWS, IS_MAC

def _open_path(path: str):
    """Open a file/folder in the OS default app."""
    try:
        if IS_WINDOWS:
            os.startfile(path)  # type: ignore
        elif sys.platform == "darwin":
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])
    except Exception as e:
        print(f"Failed to open path {path}: {e}")


def _maximize_window(win):
    """Maximize a Tk/CTk window cross-platform without taskbar clipping."""
    win.update_idletasks()
    system = sys.platform

    if system == "win32":
        try:
            import ctypes
            from ctypes import wintypes
            rect = wintypes.RECT()
            # SPI_GETWORKAREA = 48 (returns work area excluding taskbar)
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


def _center_modal_on_parent(win, parent):
    """Size and center modal dialog directly before display to eliminate visual jump/flicker."""
    win.withdraw()
    parent.update_idletasks()

    if sys.platform == "win32":
        try:
            import ctypes
            from ctypes import wintypes
            rect = wintypes.RECT()
            if ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0):
                ref_x, ref_y = rect.left, rect.top
                ref_w = rect.right - rect.left
                ref_h = rect.bottom - rect.top
            else:
                raise RuntimeError("SPI failed")
        except Exception:
            ref_x, ref_y = 0, 0
            ref_w = win.winfo_screenwidth()
            ref_h = win.winfo_screenheight()
    else:
        ref_x = parent.winfo_rootx()
        ref_y = parent.winfo_rooty()
        ref_w = parent.winfo_width()
        ref_h = parent.winfo_height()
        if ref_w < 400 or ref_h < 300:
            ref_x, ref_y = 0, 0
            ref_w = win.winfo_screenwidth()
            ref_h = win.winfo_screenheight()

    max_avail_w = max(700, ref_w - 30)
    max_avail_h = max(500, ref_h - 55)
    target_w = min(1180, max(920, int(ref_w * 0.86)), max_avail_w)
    target_h = min(660, max(520, int(ref_h * 0.80)), max_avail_h)

    x = ref_x + max(0, (ref_w - target_w) // 2)
    y = ref_y + max(5, (ref_h - target_h) // 2 - 15)

    win.geometry(f"{target_w}x{target_h}+{x}+{y}")
