# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['D:\Behavioral genetics_V1\Metamorph_scans\ZebraBox_Tejia\Behavior_data_070124_Pushkar\ZebraBox\Tejia_behavior_code_V1\Non-Binned_app_Zebrabox\main_window.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'pandas',
        'numpy',
        'matplotlib',
        'matplotlib.backends.backend_qt5agg',
        'PyQt5',
        'PyQt5.QtWidgets',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'cv2',
        'openpyxl',
        'pathlib',
        'PIL',
    ],
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
    name='DanioAnalyzer',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DanioAnalyzer',
)
