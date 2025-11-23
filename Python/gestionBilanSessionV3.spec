# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\Users\\vt238770\\Documents\\_CEA\\Prog\\Python\\apps\\gestionBilanSession\\gestionBilanSessionV3.py'],
    pathex=['C:\\Users\\vt238770\\Documents\\_CEA\\Prog\\Python\\apps\\gestionBilanSession\\vte'],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='gestionBilanSessionV3',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
