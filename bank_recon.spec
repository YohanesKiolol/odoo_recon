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
        (str(ROOT / 'pdf_summary_generator.py'),'.'),
        (str(ROOT / 'cloud_sync.py'),     '.'),
        (str(ROOT / 'readers'),           'readers'),
        (str(ROOT / 'assets'),            'assets'),   # includes app_icon.ico + app_icon.png
    ] + _pw_driver + _ctk_data,
    hiddenimports=[
        'main', 'config', 'reconciler', 'excel_writer', 'amount_utils', 'odoo_downloader',
        'journal_checker', 'journal_generator', 'odoo_journal_creator',
        'pdf_summary_generator', 'cloud_sync',
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

if sys.platform == 'darwin':
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name='Recon',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        console=False,
        disable_windowed_traceback=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
    )
    coll = COLLECT(
        exe,
        a.binaries,
        a.zipfiles,
        a.datas,
        strip=False,
        upx=True,
        upx_exclude=[],
        name='Recon'
    )
    app = BUNDLE(
        coll,
        name='Recon.app',
        icon=str(ROOT / 'assets' / 'app_icon.png'),
        bundle_identifier='com.eyerizz.recon',
        info_plist={
            'CFBundleName': 'Recon',
            'CFBundleDisplayName': 'Bank Reconciliation Studio',
            'CFBundleIdentifier': 'com.eyerizz.recon',
            'CFBundleVersion': '1.0.0',
            'CFBundleShortVersionString': '1.0.0',
            'NSHighResolutionCapable': True,
            'LSBackgroundOnly': False,
            'NSRequiresAquaSystemAppearance': False,
        }
    )
else:
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
        console=False,
        disable_windowed_traceback=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=str(ROOT / 'assets' / 'app_icon.ico'),
    )
