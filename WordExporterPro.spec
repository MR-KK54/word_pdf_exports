# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['D:/Kishore projects/WordPageExporterPro/run_gui.py'],
    pathex=['D:/Kishore projects/WordPageExporterPro'],
    binaries=[],
    datas=[
        ('C:/Users/Hxtreme/AppData/Local/Programs/Python/Python314/Lib/site-packages/customtkinter', 'customtkinter/'),
        ('D:/Kishore projects/WordPageExporterPro/word_exporter_pro', 'word_exporter_pro/'),
        ('D:/Kishore projects/WordPageExporterPro/static', 'static/')
    ],
    hiddenimports=['win32com.client', 'pythoncom', 'fitz', 'docx', 'customtkinter', 'PIL', 'PIL.Image', 'PIL.ImageDraw'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter.test'],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='WordExporterPro',
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
    icon=['D:/Kishore projects/WordPageExporterPro/static/icons/app_icon.ico']
)
