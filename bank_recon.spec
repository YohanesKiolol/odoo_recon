# bank_recon.spec — PyInstaller build spec
# Build with: pyinstaller bank_recon.spec

import sys
from pathlib import Path

block_cipher = None
ROOT = Path(SPECPATH)

from PyInstaller.utils.hooks import collect_all

datas = [
    # Bundle source modules (main + readers + ui) so --worker mode can import them
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
    (str(ROOT / 'ui'),                'ui'),
    (str(ROOT / 'readers'),           'readers'),
    (str(ROOT / 'assets'),            'assets'),   # includes app_icon.ico + app_icon.png + fonts
]
binaries = []
hiddenimports = [
    'main', 'config', 'reconciler', 'excel_writer', 'amount_utils', 'odoo_downloader',
    'journal_checker', 'journal_generator', 'odoo_journal_creator',
    'pdf_summary_generator', 'cloud_sync',
    'readers.odoo_reader', 'readers.bca_reader',
    'readers.mandiri_reader', 'readers.bri_reader',
    'ui', 'ui.theme', 'ui.widgets', 'ui.modals', 'ui.views', 'ui.controllers',
    'openpyxl', 'pdfplumber', 'pdfminer', 'pyzipper',
    'msoffcrypto', 'tkcalendar', 'babel.numbers',
    'customtkinter', 'darkdetect',
    'PIL', 'PIL.Image', 'PIL._imaging',
    'holidays', 'cryptography', 'cryptography.fernet',
]

ctk_ret = collect_all('customtkinter')
datas += ctk_ret[0]
binaries += ctk_ret[1]
hiddenimports += ctk_ret[2]

a = Analysis(
    [str(ROOT / 'gui.py')],
    pathex=[str(ROOT)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['playwright'],
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
