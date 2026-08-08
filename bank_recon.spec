# bank_recon.spec — PyInstaller build spec
# Build with: pyinstaller bank_recon.spec

import sys
from pathlib import Path

block_cipher = None
ROOT = Path(SPECPATH)

# Find Playwright driver directory to bundle node + cli.js
import importlib
_pw_spec = importlib.util.find_spec('playwright')
_pw_driver = []
if _pw_spec and _pw_spec.submodule_search_locations:
    _pw_root = Path(list(_pw_spec.submodule_search_locations)[0])
    _pw_driver_dir = _pw_root / 'driver'
    if _pw_driver_dir.exists():
        _pw_driver = [(str(_pw_driver_dir), str(Path('playwright') / 'driver'))]

# Find CustomTkinter assets (themes, images)
_ctk_spec = importlib.util.find_spec('customtkinter')
_ctk_data = []
if _ctk_spec and _ctk_spec.submodule_search_locations:
    _ctk_root = Path(list(_ctk_spec.submodule_search_locations)[0])
    if _ctk_root.exists():
        _ctk_data = [(str(_ctk_root), 'customtkinter')]

a = Analysis(
    [str(ROOT / 'gui.py')],
    pathex=[str(ROOT)],
    binaries=[],
    datas=[
        # Bundle source modules (main + readers) so --worker mode can import them
        (str(ROOT / 'main.py'),           '.'),
        (str(ROOT / 'config.py'),         '.'),
        (str(ROOT / 'reconciler.py'),     '.'),
        (str(ROOT / 'excel_writer.py'),   '.'),
        (str(ROOT / 'amount_utils.py'),   '.'),
        (str(ROOT / 'odoo_downloader.py'),'.'),
        (str(ROOT / 'journal_checker.py'),'.'),
        (str(ROOT / 'journal_generator.py'),'.'),
        (str(ROOT / 'odoo_journal_creator.py'),'.'),
        (str(ROOT / 'readers'),           'readers'),
        (str(ROOT / 'assets'),            'assets'),   # includes app_icon.ico + app_icon.png
    ] + _pw_driver + _ctk_data,
    hiddenimports=[
        'main', 'config', 'reconciler', 'excel_writer', 'amount_utils', 'odoo_downloader',
        'journal_checker', 'journal_generator', 'odoo_journal_creator',
        'readers.odoo_reader', 'readers.bca_reader',
        'readers.mandiri_reader', 'readers.bri_reader',
        'openpyxl', 'pdfplumber', 'pdfminer', 'pyzipper',
        'msoffcrypto', 'playwright', 'playwright.sync_api',
        'playwright._impl._driver', 'playwright._repo_version',
        'tkcalendar', 'babel.numbers',
        'customtkinter', 'darkdetect',
        'PIL', 'PIL.Image', 'PIL._imaging',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Recon',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,   # No black terminal window
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # .exe file icon shown in Windows Explorer, taskbar, and Alt+Tab.
    # The multi-resolution .ico (16/24/32/48/64/128/256 px) ensures Windows
    # picks the best size for each context automatically.
    icon=str(ROOT / 'assets' / 'app_icon.ico'),
)
