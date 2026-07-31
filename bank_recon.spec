# bank_recon.spec — PyInstaller build spec
# Build with: pyinstaller bank_recon.spec

import sys
from pathlib import Path

block_cipher = None
ROOT = Path(SPECPATH)

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
        (str(ROOT / 'readers'),           'readers'),
    ],
    hiddenimports=[
        'main', 'config', 'reconciler', 'excel_writer', 'amount_utils',
        'readers.odoo_reader', 'readers.bca_reader',
        'readers.mandiri_reader', 'readers.bri_reader',
        'openpyxl', 'pdfplumber', 'pdfminer', 'pyzipper',
        'msoffcrypto', 'dotenv',
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
    name='BankRekonsiliasi',
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
    # Windows only: set app icon if you have one
    # icon='icon.ico',
)
