# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['启动窗口二.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['uiautomation', 'win32com', 'win32api', 'win32con', 'win32event', 'win32gui', 'PIL', '亮屏进入桌面', '打卡并发消息'],
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
    name='某安自动打卡',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    uac_admin=True,
    icon=['ciga.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='某安自动打卡',
)
