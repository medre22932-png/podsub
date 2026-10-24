# -*- mode: python ; coding: utf-8 -*-
# macOS build spec for podsub.app

a = Analysis(
    ['podsub.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['docx'],
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
    [],
    exclude_binaries=True,
    name='podsub',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='podsub',
)

app = BUNDLE(
    coll,
    name='podsub.app',
    icon=None,
    bundle_identifier='com.medre.podsub',
)
