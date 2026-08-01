# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all

openpyxl_datas, openpyxl_binaries, openpyxl_hiddenimports = collect_all('openpyxl')
et_xmlfile_datas, et_xmlfile_binaries, et_xmlfile_hiddenimports = collect_all('et_xmlfile')


a = Analysis(
    ['茉莉青提计算器.py'],
    pathex=[],
    binaries=openpyxl_binaries + et_xmlfile_binaries,
    datas=openpyxl_datas + et_xmlfile_datas,
    hiddenimports=openpyxl_hiddenimports + et_xmlfile_hiddenimports,
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
    name='茉莉青提计算器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='提成计算器.icns',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='茉莉青提计算器',
)

app = BUNDLE(
    coll,
    name='茉莉青提计算器.app',
    icon='提成计算器.icns',
    bundle_identifier='com.chen.qingti-calculator',
)
