# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['../run_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['win32timezone', 'win32com.client.dynamic', 'win32com.client.gencache', 'tkinter', 'rangevalidator.a_extraction', 'rangevalidator.analysis', 'rangevalidator.b_extraction', 'rangevalidator.bitfield_catalog', 'rangevalidator.com_read', 'rangevalidator.com_values', 'rangevalidator.common', 'rangevalidator.diagnostics', 'rangevalidator.engine', 'rangevalidator.error_info', 'rangevalidator.gui', 'rangevalidator.manual_edits', 'rangevalidator.models', 'rangevalidator.preservation', 'rangevalidator.readers', 'rangevalidator.report', 'rangevalidator.result_index', 'rangevalidator.save_display', 'rangevalidator.selfcheck', 'rangevalidator.workflow', 'rangevalidator.writer'],
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
    name='RangeValidator-DirectCompare-final-Windows-x64',
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
)
