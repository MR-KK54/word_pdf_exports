"""
Word & PDF Page Exporter Pro - Windows App Builder & Installer Script
Builds standalone single-file Windows .exe application using PyInstaller.
"""

import os
import sys
import shutil
import subprocess

def build():
    print("=" * 60)
    print("Building Word & PDF Page Exporter Pro - Windows Single-File Executable...")
    print("=" * 60)

    try:
        import customtkinter
        ctk_path = customtkinter.__path__[0]
    except ImportError:
        print("Error: customtkinter module not found. Run 'pip install customtkinter'")
        sys.exit(1)

    project_root = os.path.dirname(os.path.abspath(__file__))
    ico_path = os.path.join(project_root, "static", "icons", "app_icon.ico")
    
    # Remove old build and dist folders
    shutil.rmtree(os.path.join(project_root, "build"), ignore_errors=True)
    shutil.rmtree(os.path.join(project_root, "dist"), ignore_errors=True)

    spec_content = f"""# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['{os.path.join(project_root, "run_gui.py").replace("\\", "/")}'],
    pathex=['{project_root.replace("\\", "/")}'],
    binaries=[],
    datas=[
        ('{ctk_path.replace("\\", "/")}', 'customtkinter/'),
        ('{os.path.join(project_root, "word_exporter_pro").replace("\\", "/")}', 'word_exporter_pro/'),
        ('{os.path.join(project_root, "static").replace("\\", "/")}', 'static/')
    ],
    hiddenimports=['win32com.client', 'pythoncom', 'fitz', 'docx', 'customtkinter', 'PIL', 'PIL.Image', 'PIL.ImageDraw'],
    hookspath=[],
    hooksconfig={{}},
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
    icon=['{ico_path.replace("\\", "/")}']
)
"""
    spec_file = os.path.join(project_root, "WordExporterPro.spec")
    with open(spec_file, "w", encoding="utf-8") as f:
        f.write(spec_content)

    cmd = [sys.executable, "-m", "PyInstaller", spec_file, "--noconfirm", "--clean"]

    print("Running command:", " ".join(cmd))
    res = subprocess.run(cmd, cwd=project_root)

    if res.returncode != 0:
        print("PyInstaller build failed!")
        sys.exit(res.returncode)

    exe_path = os.path.join(project_root, "dist", "WordExporterPro.exe")
    if os.path.exists(exe_path):
        size_mb = os.path.getsize(exe_path) / (1024 * 1024)
        print("\n" + "=" * 60)
        print(f"SUCCESS: Windows Single-File Executable built at:\n{exe_path} ({size_mb:.2f} MB)")
        print("=" * 60)
    else:
        print("Build finished but output executable not found.")

if __name__ == "__main__":
    build()
