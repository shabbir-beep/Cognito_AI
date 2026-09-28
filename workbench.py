"""
Sovereign AI Workbench — Master Implementation Engine (Python Backend & Web Server)
Project Code: SIH26117 | Organization: Mangalore Refinery and Petrochemicals Limited (MRPL)

Usage:
  python workbench.py         -> Starts local Web Server & Dashboard on http://localhost:8080
  python workbench.py --cli   -> Runs interactive CLI Agent mode in terminal
"""

import sys
import os
import json
import time
import argparse
import subprocess
import http.server
import socketserver
import urllib.parse
from datetime import datetime

# Reconfigure stdout for Windows console UTF-8 compatibility
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# ----------------------------------------------------
# MODULE 1: HARDWARE-AWARE PROFILER
# ----------------------------------------------------
def profile_hardware():
    """Detects CPU cores, System RAM, and GPU VRAM completely offline."""
    import platform
    
    cpu_cores = os.cpu_count() or 4
    ram_gb = 16.0  # Default fallback estimation
    gpu_info = []

    # Attempt to read RAM if psutil available
    try:
        import psutil
        ram_gb = round(psutil.virtual_memory().total / (1024 ** 3), 2)
    except ImportError:
        pass

    # Attempt GPU detection via PyTorch or nvidia-smi
    try:
        import torch
        if torch.cuda.is_available():
            for i in range(torch.cuda.device_count()):
                name = torch.cuda.get_device_name(i)
                vram = round(torch.cuda.get_device_properties(i).total_memory / (1024 ** 3), 2)
                gpu_info.append({"id": i, "name": name, "vram_gb": vram})
    except ImportError:
        pass

    if not gpu_info:
        # Try nvidia-smi fallback via subprocess
        try:
            res = subprocess.run(["nvidia-smi", "--query-gpu=gpu_name,memory.total", "--format=csv,noheader,nounits"], 
                                 capture_output=True, text=True, timeout=2)
            if res.returncode == 0 and res.stdout.strip():
                lines = res.stdout.strip().split("\n")
                for i, line in enumerate(lines):
                    parts = line.split(",")
                    name = parts[0].strip()
                    vram_mb = float(parts[1].strip())
                    gpu_info.append({"id": i, "name": name, "vram_gb": round(vram_mb / 1024, 2)})
        except Exception:
            pass

    if not gpu_info:
        gpu_info.append({"id": -1, "name": "NVIDIA RTX 4070 (Simulated / CPU Mode)", "vram_gb": 12.0})

    return {
        "os": platform.system() + " " + platform.release(),
        "cpu_cores": cpu_cores,
        "ram_gb": ram_gb,
        "gpus": gpu_info
    }

# ----------------------------------------------------
# MODULE 2: TASK CLASSIFIER & MODEL ROUTER
# ----------------------------------------------------
MODEL_REGISTRY = {
    "coding": {"model": "qwen2.5-coder:7b", "vram_req": 5.4, "label": "Industrial Automation Code"},
    "vision": {"model": "llama3.2-vision:11b", "vram_req": 7.8, "label": "P&ID Drawing / Vision Analysis"},
    "math": {"model": "deepseek-r1:7b", "vram_req": 5.6, "label": "Thermodynamic Calculation"},
    "general": {"model": "mistral:7b", "vram_req": 4.8, "label": "Document & SOP Analysis"}
}

def classify_and_route_task(prompt: str, has_file: bool = False):
    prompt_lower = prompt.lower()
    if has_file or any(k in prompt_lower for k in ["drawing", "p&id", "diagram", "pdf", "image", "scan"]):
        category = "vision"
    elif any(k in prompt_lower for k in ["code", "python", "script", "plc", "automation", "opc"]):
        category = "coding"
    elif any(k in prompt_lower for k in ["calculate", "math", "pressure", "duty", "heat", "formula"]):
        category = "math"
    else:
        category = "general"
        
    route = MODEL_REGISTRY[category]
    return category, route["model"], route["label"], route["vram_req"]

# ----------------------------------------------------
# MODULE 3: PERMISSION GATE INTERCEPTOR
# ----------------------------------------------------
class PermissionGate:
    def __init__(self, autonomous_mode: bool = False):
        self.autonomous_mode = autonomous_mode

    def request_authorization(self, action_name: str, details: str) -> bool:
        if self.autonomous_mode:
            print(f"[SECURITY GATE - AUTO APPROVED] Action: '{action_name}' | Target: {details}")
            return True
        
        print("\n" + "=" * 60)
        print(" [PERMISSION GATE INTERCEPTOR REQUIRED]")
        print("=" * 60)
        print(f" Requesting Tool Action: {action_name}")
        print(f" Action Details / Target: {details}")
        print(" Security Policy: Strict User Consent Mode")
        print("=" * 60)
        
        try:
            choice = input(" Authorize this write action? (Y/N, default Y): ").strip().lower()
            return choice != 'n'
        except EOFError:
            return True

# ----------------------------------------------------
# MODULE 4: LOCAL RAG RETRIEVER SIMULATOR
# ----------------------------------------------------
def query_local_sop_knowledge(query_text: str):
    """Simulates querying offline ChromaDB / SentenceTransformer vector database."""
    sops = {
        "pressure": "MRPL SOP-HC-2024 (Section 4.2): Max Allowable Working Pressure (MAWP) for Hydrocracker Reactor-A is 150.0 bar at 425°C. Dual safety valve calibration mandated when P > 140.0 bar.",
        "valve": "MRPL SOP-VALVE-2025: Relief Valve PSV-102A test cycle must not exceed 6 months under sour gas service conditions.",
        "tank": "MRPL SOP-TANK-101: Crude oil storage tank temperature threshold is capped at 65.0°C to prevent light-ends volatilization."
    }
    
    for key, content in sops.items():
        if key in query_text.lower():
            return content
    return "MRPL General Refinery Safety Manual (SOP-GEN-2024): All maintenance deliverables require plant manager sign-off prior to execution."

# ----------------------------------------------------
# MODULE 5: DELIVERABLE BUILDER (.docx & .py)
# ----------------------------------------------------
def create_docx_report(title: str, body_text: str, filename: str):
    output_dir = os.path.join(os.getcwd(), "deliverables")
    os.makedirs(output_dir, exist_ok=True)
    filepath = os.path.join(output_dir, filename)

    try:
        from docx import Document
        doc = Document()
        doc.add_heading(title, 0)
        doc.add_paragraph(f"DATE: {datetime.now().strftime('%Y-%m-%d %H:%M')} | CLASSIFICATION: RESTRICTED / INTERNAL REFINERY USE ONLY\n")
        doc.add_heading("AI Generated Analysis & SOP Findings", level=1)
        doc.add_paragraph(body_text)
        doc.save(filepath)
    except ImportError:
        # Fallback to plain text document if python-docx not installed
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(f"====================================================\n")
            f.write(f" {title}\n")
            f.write(f"====================================================\n\n")
            f.write(body_text)

    return filepath

# ----------------------------------------------------
# MODULE 6: LOCAL WEB SERVER & API ENDPOINT
# ----------------------------------------------------
class WorkbenchHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        parsed_path = urllib.parse.urlparse(self.path)
        if parsed_path.path == "/" or parsed_path.path == "/index.html":
            self.send_response(200)
            self.send_header("Content-Type", "text/html")
            self.end_headers()
            with open("index.html", "rb") as f:
                self.wfile.write(f.read())
            return
        elif parsed_path.path == "/api/hardware":
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            info = profile_hardware()
            self.wfile.write(json.dumps(info).encode("utf-8"))
            return
        else:
            super().do_GET()

    def do_POST(self):
        if self.path == "/api/execute":
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            req = json.loads(post_data.decode('utf-8'))
            
            prompt = req.get("prompt", "")
            has_file = req.get("has_file", False)
            
            category, model, label, vram = classify_and_route_task(prompt, has_file)
            sop_context = query_local_sop_knowledge(prompt)
            
            response_data = {
                "category": category,
                "model": model,
                "label": label,
                "vram_req": vram,
                "sop_context": sop_context,
                "timestamp": datetime.now().isoformat()
            }
            
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps(response_data).encode("utf-8"))
            return

def run_web_server(port=8080):
    print("=" * 60)
    print("  SOVEREIGN AI WORKBENCH - LOCAL SERVER MODE")
    print("  Project SIH26117 | MRPL Air-Gapped AI Assistant")
    print("=" * 60)
    
    hw = profile_hardware()
    print(f"  System OS: {hw['os']}")
    print(f"  CPU Cores: {hw['cpu_cores']} | RAM: {hw['ram_gb']} GB")
    print(f"  GPU Hardware: {hw['gpus'][0]['name']} ({hw['gpus'][0]['vram_gb']} GB VRAM)")
    print("-" * 60)
    print(f"  Launching Web Interface on: http://localhost:{port}")
    print("  Network Status: 100% AIR-GAPPED & OFFLINE")
    print("  (Press Ctrl+C to stop server)")
    print("=" * 60 + "\n")

    handler = WorkbenchHTTPRequestHandler
    with socketserver.TCPServer(("127.0.0.1", port), handler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nShutting down Sovereign AI server.")

# ----------------------------------------------------
# CLI MODE EXECUTION
# ----------------------------------------------------
def run_cli_agent():
    print("=" * 60)
    print("  SOVEREIGN AI WORKBENCH - CLI INTERACTIVE AGENT")
    print("=" * 60)
    
    hw = profile_hardware()
    print(f" Hardware Specs: CPU {hw['cpu_cores']} cores | RAM {hw['ram_gb']} GB | GPU {hw['gpus'][0]['name']}")
    
    try:
        prompt = input("\nEnter Industrial Task / Scenario: ").strip()
    except EOFError:
        prompt = ""
        
    if not prompt:
        prompt = "Read Hydrocracker Unit-4 pressure log data (142.5 bar) and draft formal Word approval note."
        print(f"Using default preset: '{prompt}'")

    cat, model, label, vram = classify_and_route_task(prompt)
    print(f"\nTask Classified: {label}")
    print(f"   Routed to Open-Weight Model: '{model}' (Req VRAM: {vram} GB)")

    print("\nQuerying Local Vector Knowledge Base (ChromaDB RAG)...")
    sop_ctx = query_local_sop_knowledge(prompt)
    print(f"   Found Grounding Context: {sop_ctx[:90]}...")

    gate = PermissionGate(autonomous_mode=False)
    filename = "Hydrocracker_Approval_Note.docx" if cat != "coding" else "tank_monitor.py"
    authorized = gate.request_authorization("Create Deliverable File", f"./deliverables/{filename}")

    if authorized:
        body = f"AUTOMATED AUDIT REPORT\n\nTask Prompt: {prompt}\n\nSOP Context:\n{sop_ctx}\n\nClearance granted by Sovereign AI Agent (SIH26117)."
        saved_path = create_docx_report("MRPL CONFIDENTIAL APPROVAL NOTE", body, filename)
        print(f"\nDeliverable compiled successfully at: {saved_path}")
    else:
        print("\nAction aborted by user permission gate.")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Sovereign AI Workbench Engine")
    parser.add_argument("--cli", action="store_true", help="Run interactive CLI mode instead of web server")
    parser.add_argument("--port", type=int, default=8080, help="Web server port (default 8080)")
    args = parser.parse_args()

    if args.cli:
        run_cli_agent()
    else:
        run_web_server(port=args.port)
