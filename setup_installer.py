"""
Sovereign AI Workbench — Standalone Windows Installer Setup
Project Code: SIH26117 | Organization: Mangalore Refinery and Petrochemicals Limited (MRPL)
"""

import sys
import os
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
    print("   👑 SOVEREIGN AI WORKBENCH - ONE-CLICK INSTALLATION WIZARD")
    print("   Project SIH26117 | MRPL Air-Gapped Industrial AI Assistant")
    print("=" * 65)
    
    workspace = os.path.dirname(os.path.abspath(__file__))
    local_app_data = os.environ.get("LOCALAPPDATA", r"C:\Users\Public")
    target_dir = os.path.join(local_app_data, "SovereignAIWorkbench")
    
    print(f"\n Installation Target Directory: {target_dir}")
    print(" Installing core assets, local RAG database, and open LLM router...")

    if os.path.exists(target_dir):
        try:
            shutil.rmtree(target_dir, ignore_errors=True)
        except Exception as e:
            print(f" Warning clearing old directory: {e}")

    os.makedirs(target_dir, exist_ok=True)

    # Copy files
    files_to_copy = ["app_ui.html", "sovereign_engine.py", "main.py", "README.md"]
    for fname in files_to_copy:
        src = os.path.join(workspace, fname)
        dst = os.path.join(target_dir, fname)
        if os.path.exists(src):
            shutil.copy2(src, dst)
            print(f"  -> Installed: {fname}")

    # Desktop Shortcut Creation via PowerShell
    desktop_path = os.path.join(os.environ.get("USERPROFILE", r"C:\Users\Public"), "Desktop", "Sovereign AI Workbench.lnk")
    python_exe = sys.executable
    main_py = os.path.join(target_dir, "main.py")

    ps_cmd = f'$s=(New-Object -COM WScript.Shell).CreateShortcut("{desktop_path}");$s.TargetPath="{python_exe}";$s.Arguments="{main_py}";$s.WorkingDirectory="{target_dir}";$s.Save()'
    try:
        subprocess.run(["powershell", "-Command", ps_cmd], capture_output=True)
        print(f"\n ✅ Created Desktop Shortcut: {desktop_path}")
    except Exception as e:
        print(f" Shortcut creation note: {e}")

    print("\n" + "=" * 65)
    print("   INSTALLATION COMPLETE!")
    print("   Launching Sovereign AI Workbench Desktop App...")
    print("=" * 65)
    
    subprocess.Popen([python_exe, main_py], cwd=target_dir)

if __name__ == '__main__':
    run_installer()
    time.sleep(2)
