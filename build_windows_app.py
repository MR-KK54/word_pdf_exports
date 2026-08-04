"""
Word & PDF Page Exporter Pro - Windows App Builder & Installer Script
Builds standalone Windows .exe application using PyInstaller and packages shortcut setup script.
"""

import os
import sys
import shutil
import subprocess

def build():
    print("=" * 60)
    print("Building Word & PDF Page Exporter Pro - Windows Executable...")
    print("=" * 60)

    # Locate customtkinter
    try:
        import customtkinter
        ctk_path = customtkinter.__path__[0]
    except ImportError:
        print("Error: customtkinter module not found. Run 'pip install customtkinter'")
        sys.exit(1)

    project_root = os.path.dirname(os.path.abspath(__file__))
    ico_path = os.path.join(project_root, "static", "icons", "app_icon.ico")
    
    # Construct PyInstaller command
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name=WordExporterPro",
        "--onefile",
        "--windowed",
        f"--icon={ico_path}" if os.path.exists(ico_path) else "",
        f"--add-data={ctk_path};customtkinter/",
        f"--add-data={os.path.join(project_root, 'word_exporter_pro')};word_exporter_pro/",
        f"--add-data={os.path.join(project_root, 'static')};static/",
        "--hidden-import=win32com.client",
        "--hidden-import=pythoncom",
        "--hidden-import=fitz",
        "--hidden-import=docx",
        "--hidden-import=customtkinter",
        os.path.join(project_root, "run_gui.py")
    ]
    # Remove empty args
    cmd = [c for c in cmd if c]

    print("Running command:", " ".join(cmd))
    res = subprocess.run(cmd, cwd=project_root)

    if res.returncode != 0:
        print("PyInstaller build failed!")
        sys.exit(res.returncode)

    exe_path = os.path.join(project_root, "dist", "WordExporterPro.exe")
    if os.path.exists(exe_path):
        print("\n" + "=" * 60)
        print(f"SUCCESS: Windows Application built at:\n{exe_path}")
        print("=" * 60)
    else:
        print("Build finished but output executable not found.")

if __name__ == "__main__":
    build()
