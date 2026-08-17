"""
build_sales.py — Cross-platform standalone builder for Sales Portal.

Usage:
  python build_sales.py

Produces:
  - Windows: dist/SalesPortal.exe (Single-file, no terminal, no external dependencies)
  - macOS:   dist/SalesPortal.app or dist/SalesPortal (Standalone macOS application)
"""

import sys
import subprocess
import shutil
from pathlib import Path


def build():
    print("=" * 60)
    print("  Building Standalone Sales Portal Application")
    print("=" * 60)

    # Check / install pyinstaller in active environment
    try:
        import PyInstaller
    except ImportError:
        print("[+] Installing pyinstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name=SalesPortal",
        "--noconfirm",
        "--clean",
        "--windowed",      # No console window
        "--onefile",       # Single standalone binary/exe
        "--collect-all", "customtkinter",
        "sales_gui.py"
    ]

    # Optional icon if exists
    icon_win = Path("assets/icon.ico")
    icon_mac = Path("assets/icon.icns")
    if sys.platform == "win32" and icon_win.exists():
        cmd.extend(["--icon", str(icon_win)])
    elif sys.platform == "darwin" and icon_mac.exists():
        cmd.extend(["--icon", str(icon_mac)])

    print("\n[+] Running PyInstaller command:")
    print("   ", " ".join(cmd))
    print()

    subprocess.check_call(cmd)

    dist_dir = Path("dist")
    print("\n" + "=" * 60)
    if sys.platform == "win32":
        exe_path = dist_dir / "SalesPortal.exe"
        if exe_path.exists():
            print(f"✅ Build SUCCESSFUL!")
            print(f"📁 Standalone Windows executable ready at:")
            print(f"   {exe_path.resolve()}")
            print(f"\n💡 Sales users can run SalesPortal.exe anywhere without Python or .env!")
    elif sys.platform == "darwin":
        app_path = dist_dir / "SalesPortal.app"
        bin_path = dist_dir / "SalesPortal"
        out = app_path if app_path.exists() else bin_path
        print(f"✅ Build SUCCESSFUL!")
        print(f"📁 Standalone macOS app ready at:")
        print(f"   {out.resolve()}")
        print(f"\n💡 Sales users can double-click to run without installing Python or .env!")
    print("=" * 60)


if __name__ == "__main__":
    build()
