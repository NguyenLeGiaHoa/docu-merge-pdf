"""
DocuMergePDF - Standalone Portable Windows Executable (.exe) Compilation Script
Optimized for ultra-compact size and fast launch time using PyInstaller.
Run: python build_exe.py
"""

import os
import sys
import subprocess
import shutil

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def check_and_install_pyinstaller():
    try:
        import PyInstaller
        print("[✓] PyInstaller is installed.")
    except ImportError:
        print("[*] Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])


def build_executable():
    check_and_install_pyinstaller()

    root_dir = os.path.dirname(os.path.abspath(__file__))
    main_script = os.path.join(root_dir, "main.py")
    dist_dir = os.path.join(root_dir, "dist")
    icon_path = os.path.join(root_dir, "app_icon.ico")

    print("[*] Starting optimized PyInstaller compilation for DocuMergePDF...")

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--clean",
        "--noconsole",
        "--onefile",
        "--name=DocuMergePDF",
        f"--icon={icon_path}",
        f"--add-data={os.path.join(root_dir, 'core')};core",
        f"--add-data={os.path.join(root_dir, 'gui')};gui",
        f"--add-data={os.path.join(root_dir, 'demo_data')};demo_data",
        f"--add-data={os.path.join(root_dir, 'TERMS_OF_SERVICE.txt')};.",
        f"--add-data={os.path.join(root_dir, 'LICENSE.txt')};.",
        f"--add-data={os.path.join(root_dir, 'README.txt')};.",
        f"--add-data={icon_path};.",
        "--collect-all=customtkinter",
        "--collect-all=reportlab",
        "--hidden-import=docx",
        "--hidden-import=openpyxl",
        "--hidden-import=pandas",
        "--hidden-import=PIL",
        "--hidden-import=fitz",
        "--hidden-import=pymupdf",
        "--hidden-import=windnd",
        "--exclude-module=matplotlib",
        "--exclude-module=scipy",
        "--exclude-module=torch",
        "--exclude-module=torchaudio",
        "--exclude-module=torchvision",
        "--exclude-module=tensorflow",
        "--exclude-module=onnxruntime",
        "--exclude-module=IPython",
        "--exclude-module=jupyter",
        "--exclude-module=notebook",
        "--exclude-module=pytest",
        main_script
    ]

    print(f"[*] Executing command:\n{' '.join(cmd)}\n")
    ret = subprocess.call(cmd, cwd=root_dir)

    if ret == 0:
        exe_src = os.path.join(dist_dir, "DocuMergePDF.exe")
        exe_root = os.path.join(root_dir, "DocuMergePDF.exe")
        if os.path.exists(exe_src):
            shutil.copy2(exe_src, exe_root)
            print(f"[✓] Successfully compiled: {exe_root}")
            print(f"[✓] File size: {os.path.getsize(exe_root) / (1024*1024):.2f} MB")
    else:
        print(f"[!] Compilation failed with exit code: {ret}")
        sys.exit(ret)


if __name__ == "__main__":
    build_executable()

