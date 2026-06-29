# -*- mode: python ; coding: utf-8 -*-
# R(판관비)·S(기타손익) 조서 생성 EXE (대화형). 빌드: pyinstaller R-S생성.spec


a = Analysis(
    ['src\\easy_run_rs.py'],
    pathex=['src'],
    binaries=[],
    datas=[('_internal/config', '_internal/config')],
    hiddenimports=['openpyxl', 'xlrd', 'yaml'],
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
    name='R-S생성',
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
