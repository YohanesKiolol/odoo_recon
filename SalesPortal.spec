# -*- mode: python ; coding: utf-8 -*-
import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_all

ROOT = Path(SPECPATH)

datas = [
    (str(ROOT / 'cloud_sync.py'), '.'),
    (str(ROOT / 'config.py'), '.'),
    (str(ROOT / 'assets'), 'assets'),
]
binaries = []
hiddenimports = [
    'cloud_sync', 'config',
    'customtkinter', 'darkdetect',
    'PIL', 'PIL.Image', 'PIL._imaging',
    'cryptography', 'cryptography.fernet',
]

tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]
binaries += tmp_ret[1]
hiddenimports += tmp_ret[2]

a = Analysis(
    ['sales_gui.py'],
    pathex=[str(ROOT)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

if sys.platform == 'darwin':
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name='SalesPortal',
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
        name='SalesPortal'
    )
    app = BUNDLE(
        coll,
        name='SalesPortal.app',
        icon=str(ROOT / 'assets' / 'sales_app_icon.png'),
        bundle_identifier='com.eyerizz.salesportal',
        info_plist={
            'CFBundleName': 'SalesPortal',
            'CFBundleDisplayName': 'Sales Discrepancy Portal',
            'CFBundleIdentifier': 'com.eyerizz.salesportal',
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
        a.datas,
        [],
        name='SalesPortal',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        upx_exclude=[],
        runtime_tmpdir=None,
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=str(ROOT / 'assets' / 'sales_app_icon.ico'),
    )
