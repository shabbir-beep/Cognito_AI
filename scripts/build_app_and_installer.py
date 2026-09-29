"""
Cognito AI — Optimized PyInstaller & Windows Installer Builder
Project Code: SIH26117 | MRPL
"""

import sys
import os
import shutil
import subprocess

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

PYINSTALLER_EXE = r"C:\Users\Stech\AppData\Roaming\Python\Python314\Scripts\pyinstaller.exe"

def build():
    workspace = os.path.abspath(os.path.dirname(__file__))
    print("=" * 60)
    print("  BUILDING COGNITO AI WORKBENCH STANDALONE EXECUTABLE")
    print("=" * 60)

    # 1. Build main application executable with heavy AI libs excluded for fast packaging
    cmd_app = [
        PYINSTALLER_EXE,
        "--noconfirm",
        "--onedir",
        "--windowed",
        "--name", "CognitoAI",
        "--exclude-module", "torch",
        "--exclude-module", "scipy",
        "--exclude-module", "numpy",
        "--exclude-module", "matplotlib",
        "--add-data", f"{os.path.join(workspace, 'app_ui.html')};.",
        "--add-data", f"{os.path.join(workspace, 'sovereign_engine.py')};.",
        os.path.join(workspace, "main.py")
    ]
    
    print("Running PyInstaller for main application...")
    res1 = subprocess.run(cmd_app, capture_output=True, text=True, cwd=workspace)
    if res1.returncode == 0:
        print("✅ Application bundle created successfully in ./dist/CognitoAI/")
    else:
        print("Application packaging output:\n", res1.stderr[-1000:])

    # 2. Build standalone single-file Installer executable setup_installer.py
    installer_script = os.path.join(workspace, "setup_installer.py")
    installer_code = '''"""
Cognito AI - One-Click Installer Setup
Project SIH26117 | MRPL
"""
import os
import sys
import shutil
import subprocess
import time

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

def run_installer():
    print("=" * 65)
    print("   COGNITO AI WORKBENCH - ONE-CLICK INSTALLATION WIZARD")
    print("   Project SIH26117 | MRPL Air-Gapped Industrial AI Assistant")
    print("=" * 65)
    
    local_app_data = os.environ.get("LOCALAPPDATA", r"C:\\Users\\Public")
    target_dir = os.path.join(local_app_data, "CognitoAI")
    
    base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    bundle_source = os.path.join(base_dir, "CognitoAI")
    if not os.path.exists(bundle_source):
        bundle_source = os.path.join(base_dir, "dist", "CognitoAI")
        
    if not os.path.exists(bundle_source):
        bundle_source = os.path.join(r"c:\\Users\\Stech\\Documents\\antigravity\\busy-lovelace", "dist", "CognitoAI")

    print(f" Target Directory: {target_dir}")
    print(" Installing core assets, local RAG database, and open LLM router...")
    
    if os.path.exists(target_dir):
        try:
            shutil.rmtree(target_dir, ignore_errors=True)
        except Exception:
            pass
            
    try:
        shutil.copytree(bundle_source, target_dir, dirs_exist_ok=True)
        print(" ✅ Files extracted and installed successfully!")
    except Exception as e:
        print(f" Note during copy: {e}")

    # Desktop Shortcut Creation via PowerShell
    target_exe = os.path.join(target_dir, "CognitoAI.exe")
    desktop_path = os.path.join(os.environ.get("USERPROFILE", r"C:\\"), "Desktop", "Cognito AI.lnk")
    
    ps_cmd = f'$s=(New-Object -COM WScript.Shell).CreateShortcut("{desktop_path}");$s.TargetPath="{target_exe}";$s.WorkingDirectory="{target_dir}";$s.Save()'
    try:
        subprocess.run(["powershell", "-Command", ps_cmd], capture_output=True)
        print(f" ✅ Created Desktop Shortcut: {desktop_path}")
    except Exception as e:
        print(f" Shortcut note: {e}")

    print("\\n" + "=" * 65)
    print("   INSTALLATION COMPLETE!")
    print("   Launching Cognito AI Desktop App...")
    print("=" * 65)
    
    if os.path.exists(target_exe):
        subprocess.Popen([target_exe], cwd=target_dir)
    else:
        print("   Starting application via main.py...")
        subprocess.Popen([sys.executable, os.path.join(bundle_source, "main.py")], cwd=bundle_source)

if __name__ == '__main__':
    run_installer()
    time.sleep(3)
'''
    with open(installer_script, "w", encoding="utf-8") as f:
        f.write(installer_code)

    cmd_installer = [
        PYINSTALLER_EXE,
        "--noconfirm",
        "--onefile",
        "--name", "Install_Cognito_AI_Workbench",
        "--exclude-module", "torch",
        "--exclude-module", "scipy",
        "--exclude-module", "numpy",
        "--add-data", f"{os.path.join(workspace, 'dist', 'CognitoAI')};CognitoAI",
        installer_script
    ]
    
    print("\nPackaging One-Click Installer Executable (Install_Cognito_AI_Workbench.exe)...")
    res2 = subprocess.run(cmd_installer, capture_output=True, text=True, cwd=workspace)
    if res2.returncode == 0:
        print("\n============================================================")
        print(" ✅ SUCCESS: ONE-CLICK INSTALLER CREATED!")
        print(" Location: ./dist/Install_Cognito_AI_Workbench.exe")
        print("============================================================")
    else:
        print("Installer build output:\n", res2.stderr[-1000:])

if __name__ == '__main__':
    build()
