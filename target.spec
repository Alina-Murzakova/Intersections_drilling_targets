# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['GUI.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\python\\Lib\\site-packages\\customtkinter', 'customtkinter/'), ('logo2.png', '.'), ('point.png', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='target',
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
    icon=['C:\\Users\\Alina\\Desktop\\Coordinates_VNS_ZBS\\logo.ico'],
)