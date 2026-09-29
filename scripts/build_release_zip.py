"""
Build Release Script for Cognito AI
Compiles the executable and packages the portable .zip distribution
including pre-made data/ directory structure.
"""

import os
import sys
import shutil
import zipfile
import subprocess

def run_command(cmd):
    print(f"Running: {cmd}")
    res = subprocess.run(cmd, shell=True)
    if res.returncode != 0:
        print(f"Error executing: {cmd}")
        sys.exit(res.returncode)

def main():
    root_dir = os.path.dirname(os.path.abspath(__file__))
    dist_dir = os.path.join(root_dir, 'dist')
    release_folder = os.path.join(dist_dir, 'CognitoAI_Portable')
    zip_path = os.path.join(root_dir, 'CognitoAI_v2.0_Portable.zip')

    pyinstaller_exe = r"C:\Users\Stech\AppData\Roaming\Python\Python314\Scripts\pyinstaller.exe"
    if not os.path.exists(pyinstaller_exe):
        pyinstaller_exe = "pyinstaller"

    # 1. Rebuild executable using PyInstaller
    print("=== 1. Compiling CognitoAI.exe with Custom Icon ===")
    cmd = (
        f'"{pyinstaller_exe}" --noconfirm --onefile --windowed '
        f'--name CognitoAI '
        f'--icon "icon.ico" '
        f'--add-data "app_ui.html;." '
        f'--add-data "sovereign_engine.py;." '
        f'--add-data "icon.ico;." '
        f'--add-data "icon.png;." '
        f'--exclude-module torch --exclude-module scipy --exclude-module numpy '
        f'--exclude-module matplotlib --exclude-module PIL --exclude-module cv2 '
        f'--exclude-module tkinter main.py'
    )
    run_command(cmd)

    exe_src = os.path.join(dist_dir, 'CognitoAI.exe')
    if not os.path.exists(exe_src):
        print("Build failed: CognitoAI.exe not found.")
        sys.exit(1)

    # Copy binary to root workspace
    shutil.copy2(exe_src, os.path.join(root_dir, 'CognitoAI.exe'))

    # 2. Assemble Portable Release Directory
    print("=== 2. Creating Portable Distribution Folder ===")
    if os.path.exists(release_folder):
        shutil.rmtree(release_folder)

    os.makedirs(release_folder, exist_ok=True)
    shutil.copy2(exe_src, os.path.join(release_folder, 'CognitoAI.exe'))

    # Pre-made data directories
    data_dir = os.path.join(release_folder, 'data')
    os.makedirs(os.path.join(data_dir, 'settings'), exist_ok=True)
    os.makedirs(os.path.join(data_dir, 'models'), exist_ok=True)
    os.makedirs(os.path.join(data_dir, 'sessions'), exist_ok=True)
    os.makedirs(os.path.join(data_dir, 'ollama'), exist_ok=True)
    os.makedirs(os.path.join(data_dir, 'exports'), exist_ok=True)

    readme_content = """Cognito AI v2.0 — Portable Release
===================================================

Quick Start:
1. Double click CognitoAI.exe to launch.
2. All settings, models, sessions, and engine data are saved inside the 'data/' folder.
3. Keep the 'data/' folder alongside CognitoAI.exe when moving the application.

Requirements:
- Windows 10/11 64-bit or Linux
- No administrative rights or installation required.
"""
    with open(os.path.join(release_folder, 'README.txt'), 'w', encoding='utf-8') as f:
        f.write(readme_content)

    # 3. Create .zip Archive
    print("=== 3. Packaging .zip Distribution Archive ===")
    if os.path.exists(zip_path):
        os.remove(zip_path)

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(release_folder):
            for file in files:
                abs_file = os.path.join(root, file)
                rel_path = os.path.relpath(abs_file, release_folder)
                zf.write(abs_file, arcname=os.path.join('CognitoAI_Portable', rel_path))

    print("=" * 65)
    print(f" SUCCESS: Portable zip package created at:")
    print(f" {zip_path}")
    print(f" File Size: {os.path.getsize(zip_path) / (1024*1024):.2f} MB")
    print("=" * 65)

if __name__ == '__main__':
    main()
